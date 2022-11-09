VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteRecaudacionDiaria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SstKardex 
      Height          =   9105
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   16060
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RECAUDACION"
      TabPicture(0)   =   "FrmReporteRecaudacionDiaria.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTotalDevoluciones"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblnumdevoluciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblPrecioCompra"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblNumanuladas"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTotalAnuladas"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblNumEfectivo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTotalEfectivo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LblnumMastercard"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblTotalMastercard"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblNumVisa"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblTotalVisa"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblNumSubtotal"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblSubtotal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblNumTikets"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblTotalgeneral"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblnumTotal"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblTotalCreditos"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblnumCreditos"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label6(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label6(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label6(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label6(3)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblNumVitepay"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblTotalVitepay"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label6(4)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label6(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Shape3"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbltotalletras"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "LblCantidad"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Image1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdarqueodetallado"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "DtcMoneda"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdReporteDetallado"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdConsolidadoTicket"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmddetalladoo"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdproduccion"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdchasis"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "DtcVentanilla"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdVentasVendedor"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdListadorecibos"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cmdConsolidado"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "cmdsalir"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cmdImprimir"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "DtcTurno"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "SSTab1"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "HfgEfectivo"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "DtcAlmacen"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "DtpHasta"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "DtpDesde"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "DtcOperador"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "chkTurno"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "TxtBusquedaRapida"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "ChkAlmacen"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "chkOperador"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdAceptar"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "chk_ventanilla"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "chk_moneda"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "frmarqueo"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).ControlCount=   65
      TabCaption(1)   =   "REPORTE"
      TabPicture(1)   =   "FrmReporteRecaudacionDiaria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Image4"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Image2"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "Image3"
      Tab(1).Control(6)=   "Shape5"
      Tab(1).Control(7)=   "Shape6"
      Tab(1).Control(8)=   "Shape7"
      Tab(1).Control(9)=   "Chart"
      Tab(1).Control(10)=   "HfReporte"
      Tab(1).Control(11)=   "Chart2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "ESTADISTICA"
      TabPicture(2)   =   "FrmReporteRecaudacionDiaria.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame frmarqueo 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   5250
         Left            =   11160
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   6615
         Begin VB.TextBox txtsobrante_faltante 
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
            Left            =   5280
            TabIndex        =   83
            Top             =   4200
            Width           =   1095
         End
         Begin VitekeySoft.ChameleonBtn cmdConsolidadoDetallado 
            Height          =   495
            Left            =   240
            TabIndex        =   81
            Top             =   4320
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            BTYPE           =   3
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteRecaudacionDiaria.frx":0054
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txttotalsistema 
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
            Left            =   5280
            TabIndex        =   77
            Top             =   4635
            Width           =   1095
         End
         Begin VB.TextBox txt_idarqueo 
            Height          =   285
            Left            =   2400
            TabIndex        =   76
            Top             =   60
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame frmcantidad 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   4800
            TabIndex        =   73
            Top             =   1920
            Visible         =   0   'False
            Width           =   1455
            Begin VB.TextBox txtcantidad 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   330
               Left            =   720
               TabIndex        =   74
               Top             =   80
               Width           =   615
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CANT:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   120
               TabIndex        =   75
               Top             =   120
               Width           =   465
            End
         End
         Begin VB.CommandButton cmdcerrararqueo 
            Height          =   255
            Left            =   6340
            Picture         =   "FrmReporteRecaudacionDiaria.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   40
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfArqueo 
            Height          =   3735
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   6588
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
         Begin VB.Label lblfaltantesobrante 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4380
            TabIndex        =   84
            Top             =   4245
            Width           =   45
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL SISTEMA :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3975
            TabIndex        =   80
            Top             =   4680
            Width           =   1155
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RECAUDACION EFECTIVO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   330
            TabIndex        =   79
            Top             =   120
            Width           =   1725
         End
      End
      Begin VB.CheckBox chk_moneda 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "MONEDA:"
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
         Height          =   260
         Left            =   6360
         TabIndex        =   69
         Top             =   2800
         Width           =   975
      End
      Begin VB.CheckBox chk_ventanilla 
         Appearance      =   0  'Flat
         Caption         =   "VENTANILLA"
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
         Left            =   240
         TabIndex        =   62
         Top             =   2320
         Width           =   1335
      End
      Begin VitekeySoft.ChameleonBtn cmdAceptar 
         Height          =   400
         Left            =   17160
         TabIndex        =   56
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "VISUALIZAR EN PANTALLA"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":2F14
         PICN            =   "FrmReporteRecaudacionDiaria.frx":2F30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkOperador 
         Appearance      =   0  'Flat
         Caption         =   "OPERADOR"
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
         Left            =   225
         TabIndex        =   5
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CheckBox ChkAlmacen 
         Appearance      =   0  'Flat
         Caption         =   "ENTIDAD"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1905
         Width           =   1095
      End
      Begin VB.TextBox TxtBusquedaRapida 
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
         Left            =   6480
         TabIndex        =   3
         Top             =   1500
         Width           =   1575
      End
      Begin VB.CheckBox chkTurno 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TURNO :"
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
         Left            =   3600
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
      Begin MSChart20Lib.MSChart Chart2 
         Height          =   4215
         Left            =   -74640
         OleObjectBlob   =   "FrmReporteRecaudacionDiaria.frx":545B
         TabIndex        =   2
         Top             =   4560
         Width           =   9615
      End
      Begin MSDataListLib.DataCombo DtcOperador 
         Height          =   330
         Left            =   1590
         TabIndex        =   6
         Top             =   1500
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   ""
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
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   795
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   184221697
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   795
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   184221697
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   1605
         TabIndex        =   9
         Top             =   1905
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   ""
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgEfectivo 
         Height          =   5775
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   10186
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfReporte 
         Height          =   3135
         Left            =   -74640
         TabIndex        =   11
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5530
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   10080
         TabIndex        =   12
         Top             =   5280
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6588
         _Version        =   393216
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "MASTERCARD"
         TabPicture(0)   =   "FrmReporteRecaudacionDiaria.frx":8D66
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "HfDebito"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "VISA"
         TabPicture(1)   =   "FrmReporteRecaudacionDiaria.frx":8D82
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "HfCredito"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "PAGO CREDITO"
         TabPicture(2)   =   "FrmReporteRecaudacionDiaria.frx":8D9E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "HfDeudas"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "VITEPAY"
         TabPicture(3)   =   "FrmReporteRecaudacionDiaria.frx":8DBA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "HfVitepay"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "RECIBOS"
         TabPicture(4)   =   "FrmReporteRecaudacionDiaria.frx":8DD6
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "HfDevoluciones"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "DEPOSITOS"
         TabPicture(5)   =   "FrmReporteRecaudacionDiaria.frx":8DF2
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "HfDeposito"
         Tab(5).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDebito 
            Height          =   3135
            Left            =   240
            TabIndex        =   13
            Top             =   540
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5530
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCredito 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   14
            Top             =   540
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5530
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDeudas 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   15
            Top             =   540
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   5530
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVitepay 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   16
            Top             =   540
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDevoluciones 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   17
            Top             =   540
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5530
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDeposito 
            Height          =   3135
            Left            =   -74760
            TabIndex        =   18
            Top             =   540
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5530
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
      End
      Begin MSChart20Lib.MSChart Chart 
         Height          =   7935
         Left            =   -64920
         OleObjectBlob   =   "FrmReporteRecaudacionDiaria.frx":8E0E
         TabIndex        =   19
         Top             =   840
         Width           =   9615
      End
      Begin MSDataListLib.DataCombo DtcTurno 
         Height          =   330
         Left            =   4800
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VitekeySoft.ChameleonBtn cmdImprimir 
         Height          =   400
         Left            =   17160
         TabIndex        =   57
         Top             =   920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "REPORTE DETALLADO [A4]  "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":C719
         PICN            =   "FrmReporteRecaudacionDiaria.frx":C735
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsalir 
         Height          =   405
         Left            =   17160
         TabIndex        =   58
         Top             =   4320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "CERRAR                               "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":EA89
         PICN            =   "FrmReporteRecaudacionDiaria.frx":EAA5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdConsolidado 
         Height          =   405
         Left            =   10560
         TabIndex        =   59
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "REPORTE CONSOLIDADO    "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":11BDC
         PICN            =   "FrmReporteRecaudacionDiaria.frx":11BF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdListadorecibos 
         Height          =   405
         Left            =   17160
         TabIndex        =   60
         Top             =   3015
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "LISTADO DE RECIBOS            "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":13F4C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":13F68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdVentasVendedor 
         Height          =   405
         Left            =   17160
         TabIndex        =   61
         Top             =   3465
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "VENTAS X VENDEDOR          "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":162BC
         PICN            =   "FrmReporteRecaudacionDiaria.frx":162D8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcVentanilla 
         Height          =   330
         Left            =   1605
         TabIndex        =   63
         Top             =   2320
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VitekeySoft.ChameleonBtn cmdchasis 
         Height          =   405
         Left            =   17160
         TabIndex        =   64
         Top             =   3885
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "SERIES-CHASIS                     "
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
         MICON           =   "FrmReporteRecaudacionDiaria.frx":1862C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":18648
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdproduccion 
         Height          =   405
         Left            =   17160
         TabIndex        =   65
         Top             =   2220
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "REPORTE PRODUCCION      "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":1A99C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":1A9B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmddetalladoo 
         Height          =   345
         Left            =   17160
         TabIndex        =   66
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   " DETALLADO    [TICKET]      "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":1CD0C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":1CD28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdConsolidadoTicket 
         Height          =   405
         Left            =   17160
         TabIndex        =   67
         Top             =   1780
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "CONSOLIDADO [TICKET]      "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":1F07C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":1F098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdReporteDetallado 
         Height          =   405
         Left            =   10560
         TabIndex        =   68
         Top             =   4260
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "REPORTE DETALLADO          "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":213EC
         PICN            =   "FrmReporteRecaudacionDiaria.frx":21408
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   315
         Left            =   7365
         TabIndex        =   70
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VitekeySoft.ChameleonBtn cmdarqueodetallado 
         Height          =   405
         Left            =   17160
         TabIndex        =   82
         Top             =   1350
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "ARQUEO DETALLADO            "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRecaudacionDiaria.frx":2375C
         PICN            =   "FrmReporteRecaudacionDiaria.frx":23778
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   300
         Left            =   -64920
         Top             =   480
         Width           =   9615
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   300
         Left            =   -74640
         Top             =   4200
         Width           =   9615
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   300
         Left            =   -74640
         Top             =   480
         Width           =   9615
      End
      Begin VB.Image Image3 
         Height          =   270
         Left            =   -74520
         Picture         =   "FrmReporteRecaudacionDiaria.frx":25ACC
         Stretch         =   -1  'True
         Top             =   4215
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTE DETALLADO "
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
         Left            =   -73275
         TabIndex        =   55
         Top             =   525
         Width           =   1845
      End
      Begin VB.Image Image2 
         Height          =   270
         Left            =   -74520
         Picture         =   "FrmReporteRecaudacionDiaria.frx":29009
         Stretch         =   -1  'True
         Top             =   495
         Width           =   285
      End
      Begin VB.Image Image1 
         Height          =   270
         Left            =   360
         Picture         =   "FrmReporteRecaudacionDiaria.frx":2C546
         Stretch         =   -1  'True
         Top             =   2775
         Width           =   285
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECAUDACION EFECTIVO"
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
         Left            =   810
         TabIndex        =   54
         Top             =   2805
         Width           =   1995
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1695
         TabIndex        =   53
         Top             =   840
         Width           =   225
      End
      Begin VB.Label lbltotalletras 
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
         Height          =   300
         Left            =   10200
         TabIndex        =   52
         Top             =   4800
         Width           =   9375
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   735
         Left            =   15000
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL  :"
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
         Index           =   5
         Left            =   11280
         TabIndex        =   51
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SUB TOTAL :"
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
         Index           =   4
         Left            =   10965
         TabIndex        =   50
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label lblTotalVitepay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   49
         Top             =   2280
         Width           =   1785
      End
      Begin VB.Label lblNumVitepay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   48
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "VITEPAY :"
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
         Index           =   3
         Left            =   11190
         TabIndex        =   47
         Top             =   2325
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO CREDITOS :"
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
         Index           =   2
         Left            =   10575
         TabIndex        =   46
         Top             =   1965
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº TIKETS :"
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
         Left            =   15360
         TabIndex        =   45
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "VISA :"
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
         Index           =   1
         Left            =   11460
         TabIndex        =   44
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MARTER CARD :"
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
         Left            =   10695
         TabIndex        =   43
         Top             =   1245
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "EFECTIVO :"
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
         Index           =   0
         Left            =   11085
         TabIndex        =   42
         Top             =   885
         Width           =   810
      End
      Begin VB.Label lblnumCreditos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   41
         Top             =   1920
         Width           =   945
      End
      Begin VB.Label lblTotalCreditos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   40
         Top             =   1920
         Width           =   1785
      End
      Begin VB.Label lblnumTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   39
         Top             =   3315
         Width           =   945
      End
      Begin VB.Label lblTotalgeneral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   38
         Top             =   3315
         Width           =   1785
      End
      Begin VB.Label lblNumTikets 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   15240
         TabIndex        =   37
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   36
         Top             =   2640
         Width           =   1785
      End
      Begin VB.Label lblNumSubtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   35
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label lblTotalVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   34
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label lblNumVisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   33
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblTotalMastercard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   32
         Top             =   1200
         Width           =   1785
      End
      Begin VB.Label LblnumMastercard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   31
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lblTotalEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   30
         Top             =   840
         Width           =   1785
      End
      Begin VB.Label lblNumEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   29
         Top             =   840
         Width           =   945
      End
      Begin VB.Label lblTotalAnuladas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   28
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lblNumanuladas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   27
         Top             =   480
         Width           =   945
      End
      Begin VB.Label LblPrecioCompra 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº ANULADAS :"
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
         Left            =   10725
         TabIndex        =   26
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RENDIMIENTO PACIENTES"
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
         Left            =   -73530
         TabIndex        =   25
         Top             =   4245
         Width           =   2115
      End
      Begin VB.Image Image4 
         Height          =   270
         Left            =   -64800
         Picture         =   "FrmReporteRecaudacionDiaria.frx":2FA83
         Stretch         =   -1  'True
         Top             =   495
         Width           =   285
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RENDIMIENTO ACUMULADO"
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
         Left            =   -63900
         TabIndex        =   24
         Top             =   525
         Width           =   2295
      End
      Begin VB.Label lblnumdevoluciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12120
         TabIndex        =   23
         Top             =   2985
         Width           =   945
      End
      Begin VB.Label lblTotalDevoluciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   13080
         TabIndex        =   22
         Top             =   2985
         Width           =   1785
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "DEVOLUCIONES :"
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
         Index           =   6
         Left            =   10620
         TabIndex        =   21
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   120
         Top             =   480
         Width           =   8415
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   4830
         Left            =   10080
         Top             =   360
         Width           =   9735
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   300
         Left            =   240
         Top             =   2760
         Width           =   5895
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5400
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":32FC0
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":33414
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":33734
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":33B88
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":33FDC
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":342FC
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":34750
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":348AC
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":34D00
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":3501C
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":358F8
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":35C18
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRecaudacionDiaria.frx":35F38
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteRecaudacionDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn3_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub chkAlmacen_Click()
If Me.ChkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If

End Sub

Private Sub chkOperador_Click()
If Me.chkOperador.Value = 1 Then
    Me.DtcOperador.Enabled = True
Else
    Me.DtcOperador.Enabled = False
End If
End Sub

Private Sub CmdDesacer_Click()

End Sub



Private Sub chkturno_Click()
If Me.chkturno.Value = 1 Then
    Me.DtcTurno.Visible = True
    'Me.lblHorario.Visible = True
Else
    Me.DtcTurno.Visible = False
    'Me.lblHorario.Visible = False
End If
End Sub

Private Sub ClbAcciones_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub cmdAceptar_Click()
        Dim Ans As Boolean
        Dim i As Integer
        Dim Anulado As String
        Dim X As Integer
        Dim tanuladas As Double
        Dim tdeudas As Double
        Dim tcredito As Double
        Dim tdebito As Double
        Dim tefectivo As Double
        Dim pendientes As Integer
        Dim tsubtotal As Double
        Dim tTotal As Double
        StrAlmacen = ""
        StrOperador = ""
        strventanilla = ""
        lblTotalMastercard.Caption = 0
        DteInicio = CVDate(Me.DtpDesde.Value)
        DteFin = CVDate(Me.DtpHasta.Value)
        
        If Me.ChkAlmacen.Value = 1 Then
            StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
        
        End If
        
        If Me.chkOperador.Value = 1 Then
         StrOperador = Replace(Me.DtcOperador.BoundText, "'", "''")
        End If
        If Me.chk_ventanilla.Value = 1 Then
         strventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
        End If
        
        Call Llenar_efectivo(Me.HfgEfectivo, StrOperador, StrAlmacen, strventanilla)
        Call Llenar_reporte(Me.HfReporte, StrOperador, StrAlmacen, strventanilla)
        Call llenar_devoluciones(Me.HfDevoluciones, StrOperador, StrAlmacen, strventanilla)
        Call llenar_deposito_cuenta(Me.HfDeposito, StrOperador, StrAlmacen, strventanilla)
        Call Llenar_mastercard(Me.HfDebito, StrOperador, StrAlmacen, strventanilla)
        Call Llenar_visa(Me.HfCredito, StrOperador, StrAlmacen, strventanilla)
        Call llenar_anuladas(StrOperador, StrAlmacen, strventanilla)
        
        Call Llenar_pago_credito(Me.HfDeudas, StrOperador, StrAlmacen, strventanilla)
        
        Me.lblNumSubtotal.Caption = Val(Me.lblNumanuladas.Caption) + Val(Me.lblNumEfectivo.Caption) + Val(Me.LblnumMastercard.Caption) + Val(Me.lblNumVisa.Caption)
        Me.lblnumTotal.Caption = Val(Me.lblNumSubtotal.Caption) - Val(Me.lblNumanuladas.Caption)
        tsubtotal = Val(Me.lblTotalEfectivo.Caption) + Val(Me.lblTotalCreditos.Caption) + Val(Me.lblTotalMastercard.Caption) + Val(Me.lblTotalVisa.Caption) + Val(Me.lblTotalVitepay.Caption)
        tTotal = Val(tsubtotal) + Val(Me.lblTotalDevoluciones.Caption)
        Me.lblSubtotal.Caption = Format(tsubtotal, "###0.00")
        Me.lblTotalgeneral.Caption = Format(tTotal, "###0.00")
        Me.lblNumTikets.Caption = Val(Me.lblNumanuladas.Caption) + Val(Me.lblNumEfectivo.Caption) + Val(Me.LblnumMastercard.Caption) + Val(Me.lblNumVisa.Caption)
        If Val(Me.lblTotalgeneral.Caption) > 0 Then
            Me.LblTotalLetras.Caption = UCase(EnLetras(Val(lblTotalgeneral.Caption)))
        End If
End Sub

Private Sub cmdarqueodetallado_Click()
    Dim turno As String
    Dim StrOperador As String

        turno = ""
        id_alm = ""
        StrOperador = ""
        
        If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
        End If
        
        If Me.ChkAlmacen.Value = 1 Then
            id_alm = Replace(Me.DtcAlmacen.BoundText, "'", "''")
        End If
        If Me.chkOperador.Value = 1 Then
            StrOperador = Replace(Me.DtcOperador.BoundText, "'", "''")
        End If
    
        If StrOperador = "" Then
           StrOperador = KEY_USUARIO
        End If
        
        Call get_arqueo(StrOperador)
       
        
        
End Sub
Private Sub get_arqueo(ByVal in_dni As String)
Dim in_arqueo As String
strCadena = "SELECT * FROM arqueo_caja WHERE id_fecha='" & KEY_FECHA & "'  and dni_vendedor='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'  ORDER BY id_arqueo DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   in_arqueo = put_arqueo(in_dni)
   Me.txt_idarqueo.Text = in_arqueo
   Call llenar_arqueo(Me.HfArqueo, in_arqueo)
   
Else
   Me.txt_idarqueo.Text = rst("id_arqueo")
   Call llenar_arqueo(Me.HfArqueo, rst("id_arqueo"))
End If
Me.frmarqueo.Visible = True
End Sub

Public Sub llenar_arqueo(ByVal Grilla As MSHFlexGrid, ByVal in_arqueo As String)
On Error GoTo salir

strCadena = "SELECT * FROM view_arqueo WHERE id_arqueo='" & in_arqueo & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 3500
            Grilla.ColWidth(2) = 1000
            Grilla.ColWidth(3) = 1200
       Next
        cabecera = "IDBILLETE" & vbTab & "DESCRIPCION" & vbTab & "CANT" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        in_total = 0
        For i = 0 To rst.RecordCount - 1
            in_unidad = rst("cantidad") * rst("valor")
            Fila = rst("id") & vbTab & rst("billete") & vbTab & rst("cantidad") & vbTab & Format(in_unidad, "###0.00")
            Grilla.AddItem Fila
            in_total = in_total + in_unidad
            rst.MoveNext
        Next i
       
       Fila = "" & vbTab & ":::::::::::::::::TOTAL FISICO::::::::::::::::::" & vbTab & "" & vbTab & Format(in_total, "###0.00")
       Grilla.AddItem Fila
       For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H8080FF
       Next k
      
      If Val(Me.txttotalsistema.Text) > Val(in_total) Then
         Me.txtsobrante_faltante.Text = Format(Val(Me.txttotalsistema.Text) - Val(in_total), "###0.00")
         Me.lblfaltantesobrante.Caption = "FALTANTE :"
      End If
      
      If Val(Me.txttotalsistema.Text) < Val(in_total) Then
         Me.txtsobrante_faltante.Text = Format(Val(in_total) - Val(Me.txttotalsistema.Text), "###0.00")
         Me.lblfaltantesobrante.Caption = "SOBRANTE :"
      End If
      
      
      
      
      
      
     
     
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"

End Sub

Private Function put_arqueo(ByVal in_dni As String) As String

strCadena = "INSERT INTO arqueo_caja(`id_fecha`,`dni_vendedor`,`dni_supervisor`,`ruc`)VALUES('" & KEY_FECHA & "','" & in_dni & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call Execute_Sql(strCadena)
strCadena = "SELECT * FROM arqueo_caja WHERE dni_vendedor='" & in_dni & "' ORDER BY id_arqueo DESC LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   put_arqueo = rstK("id_arqueo")
   Call genera_arqueo(in_dni, put_arqueo)
End If

End Function
Private Sub genera_arqueo(ByVal in_dni As String, ByVal in_arqueo As String)
strCadena = "SELECT * FROM arqueo_billete  ORDER BY id_billete DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO arqueo_detalle(`id_arqueo`,`id_billete`) VALUES ('" & Val(in_arqueo) & "','" & rst("id_billete") & "')"
       Call Execute_Sql(strCadena)
       rst.MoveNext
   Next i
End If

End Sub

Private Sub cmdcerrararqueo_Click()
frmarqueo.Visible = False
End Sub

Private Sub cmdchasis_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value
param = arr()
          
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If

strCadena = "SELECT id_venta,`documento`,`fecha_emision`,`id_cliente`,`ncliente`,`detalle`,`descripcion`,`nro_chasis`,`serie`,`total`,`id_alm`,`nota`,`monto_nota`,`id_vendedor`,`vendedor`,`ruc` " & _
" FROM view_reporte_comisionable_chasis WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_venta_comisionable", , App.Path + "\Reportes\")
End Sub

Private Sub cmdConsolidado_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"



arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value


param = arr()

          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          in_ventanilla = ""
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          If Me.chk_ventanilla.Value = 1 Then
            in_ventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
          End If
          
          
          
          

strCadena = "SELECT id_doc,doc_des,descripcion,sum(total) FROM view_consolida_cobranza_v2 WHERE anulado='no' and  id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'  AND ruc='" & KEY_RUC & "' and id_doc IN('0001','0003','0007','0008') GROUP BY id_doc,id_forma_pago"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptResumenCobranza", param, App.Path + "\Reportes\")

End Sub
Private Sub consolidado_cobranza()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"



arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value


param = arr()

          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          in_ventanilla = ""
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          If Me.chk_ventanilla.Value = 1 Then
            in_ventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
          End If
          
          
          
          

strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja " & _
"  FROM view_reporte_detallado_ultimate WHERE  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'  AND ruc='" & KEY_RUC & "'  order by fecha_emision asc,id_doc,serie,numero ASC"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_detallado_caja_consolidado", param, App.Path + "\Reportes\")
End Sub


Private Sub cmdConsolidadoDetallado_Click()
Call impresion_consolidado_arqueo_det(Val(Me.txt_idarqueo.Text), Val(Me.txttotalsistema.Text))
End Sub

Private Sub cmdConsolidadoTicket_Click()
        
        
        If Me.ChkAlmacen.Value = 1 Then
            StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
        End If
        
        If Me.chkOperador.Value = 1 Then
         StrOperador = Replace(Me.DtcOperador.BoundText, "'", "''")
        End If
        If Me.chk_ventanilla.Value = 1 Then
            strventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
        End If
        
        
        
        Call impresion_consolidado_arqueo(Me.DtpDesde.Value, Me.DtpHasta.Value, strventanilla, StrOperador, StrAlmacen)
End Sub

Private Sub cmddetalladoo_Click()
Dim in_ni  As String
Dim in_fi As String
in_ni = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
in_fi = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
strCadena = "SELECT id_vendedor,nombre_completo,fecha_emision,documento,ncliente,id_alm,almacen,cantidad,detalle,descripcion,nro_motor,nro_chasis,id_forma_pago," & _
"id_linea,linea,total,monto_nota,nota,'" & in_ni & "','" & in_fi & "',contado,credito,cuotas,interes,ruc FROM view_produccion_v3 WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

'Ans = ShowMultiReport(rst, "rpt_produccion_detalle", , App.Path + "\Reportes\")
        turno = ""
        If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          in_ventanilla = ""
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          If Me.chk_ventanilla.Value = 1 Then
            in_ventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
          End If
          
If KEY_GRIFO = "si" Then
    Call impresion_consolidado_ticket(operador, turno, Me.DtpDesde.Value, Me.DtpHasta.Value, StrAlmacen)
Else
    Call impresion_consolidado_ticket_normal(operador, turno, Me.DtpDesde.Value, Me.DtpHasta.Value, StrAlmacen)
End If

End Sub

Private Sub cmdImprimir_Click()
Dim param As Variant
Dim cam3(0 To 2, 1 To 2)  As String


                    cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    cam3(0, 2) = Format(DtpDesde.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(DtpHasta.Value, "dd-mm-YYYY")
                    cam3(2, 2) = Trim(Me.DtcAlmacen.Text)
                    param = cam3()

          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          in_ventanilla = ""
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          If Me.chk_ventanilla.Value = 1 Then
            in_ventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
          End If
          
          
          
          
If Me.chk_moneda.Value = 1 Then
    strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja " & _
    "  FROM view_reporte_detallado_ultimate WHERE id_moneda='" & Me.DtcMoneda.BoundText & "' and  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_recibo=0 AND ruc='" & KEY_RUC & "' order by fecha_emision asc,id_doc,serie,numero ASC"
Else
   ' strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,(monto_caja-if(id_doc='0007',function_get_nota_credito_pago(id_venta,ruc)*-1,function_get_nota_credito(id_venta)) ) " & _
   "  FROM view_reporte_detallado_ultimate WHERE  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_recibo=0 AND ruc='" & KEY_RUC & "' order by id_forma_pago, fecha_emision asc,id_doc ASC,serie ASC,numero ASC"
    
    strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja " & _
   "  FROM view_reporte_detallado_ultimate WHERE  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_recibo=0 AND ruc='" & KEY_RUC & "' order by id_forma_pago, fecha_emision asc,id_doc ASC,serie ASC,numero ASC"
   
End If
Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "Rpt_detallado_caja_ii", param, App.Path + "\Reportes\")


End Sub

Private Sub cmdListadorecibos_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"



arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value


param = arr()

          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          
          
          
strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,id_forma_pago,anulado,monto,total " & _
"  FROM view_reporte_detallado_v2 WHERE id_doc='0054' and  id_vendedor LIKE '%" & Trim(StrOperador) & "%' AND  id_alm like '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' order by fecha_emision asc,id_doc,serie,numero ASC"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_detallado_caja", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdproduccion_Click()
Dim in_ni As Date
Dim in_fi As Date
in_ni = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
in_fi = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
strCadena = "SELECT id_vendedor,nombre_completo,fecha_emision,documento,id_alm,almacen,cantidad,detalle,id_linea,linea,total,monto_nota,item_nota," & CVDate(in_ni) & ", " & CVDate(in_fi) & ",ruc FROM view_produccion_v2 WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
strCadena = "select linea, sum(cantidad-item_nota) as cantidad, sum(total-monto_nota) as monto from view_produccion_v2 v where fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' group by  v.`id_linea` "
Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rst, "rpt_produccion", , App.Path + "\Reportes\", , , , , rstK, "rpt_produccion_acumulado")
   

End Sub

Private Sub cmdReporteDetallado_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"



arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value


param = arr()

          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          in_ventanilla = ""
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If
          If Me.chk_ventanilla.Value = 1 Then
            in_ventanilla = Replace(Me.DtcVentanilla.BoundText, "'", "''")
          End If
          
          
          
          

strCadena = "SELECT 0,fecha_emision,id_doc,doc_des,documento,ncliente,total,id_forma_pago,descripcion,dni_save FROM view_consolida_cobranza_v2 WHERE anulado='no' and  id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'  AND ruc='" & KEY_RUC & "' and id_doc IN('0001','0003','0007','0008')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptResumenCobranza_detallado", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVentasVendedor_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(0, 2) = Me.DtpDesde.Value
arr(1, 2) = Me.DtpHasta.Value
param = arr()
          turno = ""
          If Me.chkturno.Value = 1 Then
            turno = Replace(Me.DtcTurno.BoundText, "'", "''")
          End If
          If Me.ChkAlmacen.Value = 1 Then
             StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
          Else
              StrAlmacen = ""
          End If
       
          operador = ""
          
          If Me.chkOperador.Value = 1 Then
            operador = Replace(Me.DtcOperador.BoundText, "'", "''")
          End If

strCadena = "SELECT `id_vendedor`,`documento`,`fecha_emision`,`ncliente`,`total`,`function_get_nota_fecha`(`id_venta`,'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "') AS `nota`,`function_get_nota_credito_fecha`(`id_venta`,'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "') AS `monto_nota`,`nombre_completo`,`ruc` " & _
" FROM view_venta_vendedor_fecha WHERE  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_venta_vendedor", , App.Path + "\Reportes\")


End Sub



Private Sub Command1_Click()


End Sub

Private Sub DtcTurno_Change()
strCadena = "SELECT * FROM turno WHERE id_turno='" & Trim(Me.DtcTurno.BoundText) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   ' Me.lblHorario.Caption = "[ " & Format(rstT("hora_inicio"), "Medium Time") & Space(2) & "-" & Space(2) & Format(rstT("hora_final"), "Medium Time") & " ]"
Else
    'Me.lblHorario.Caption = ""
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim hi As String
Dim hf As String
Dim h As String
Dim strturno As String
CenterForm Me
Me.Top = 100
Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA
strCadena = "SELECT * FROM turno WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
rst.MoveFirst
strCadena = "SELECT CURTIME()"
Call ConfiguraRstT(strCadena)
h = Format(rstT(0), "hh:mm:ss")
For i = 0 To rst.RecordCount - 1
    hi = Format(rst("hora_inicio"), "hh:mm:ss")
    hf = Format(rst("hora_final"), "hh:mm:ss")
    If (Format(TimeValue(h), "hh:mm:ss") >= Format(TimeValue(hi), "hh:mm:ss")) And (Format(TimeValue(h), "hh:mm:ss") < Format(TimeValue(hf), "hh:mm:ss")) Then
        strturno = rst("id_turno")
        Exit For
    End If
        rst.MoveNext
Next i

strCadena = "SELECT id_turno as Codigo,descripcion as Descripcion FROM turno WHERE ruc='" & KEY_RUC & "' ORDER BY hora_inicio ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTurno)
Me.DtcTurno.BoundText = strturno

strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' AND E.id_medico='no' AND E.id_sucursal='" & KEY_ALM & "' AND E.id_cargo<>'00025' AND E.id_cargo<>'00015'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcOperador)
Me.DtcOperador.Enabled = False

If KEY_ALM = "00001" Then
    strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_sucursal='0' and   ruc='" & KEY_RUC & "' ORDER BY descripcion "
Else
    strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_sucursal='0' and id_tipoentidad='0' and  ruc='" & KEY_RUC & "'  ORDER BY descripcion "
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
Me.ChkAlmacen.Value = 1


strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_sucursal)) as Descripcion FROM almacen where id_sucursal='" & Me.DtcAlmacen.BoundText & "' and id_tipoentidad='00012' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcVentanilla)


strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion  FROM moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)



End Sub




Private Sub HfArqueo_Click()
If Val(Me.HfArqueo.TextMatrix(Me.HfArqueo.Row, 0)) > 0 Then
    Me.txtcantidad.Text = Me.HfArqueo.TextMatrix(Me.HfArqueo.Row, 2)
    Me.frmcantidad.Visible = True
    Call Resalta(Me.txtcantidad)
End If
End Sub

Private Sub HfgEfectivo_Click()
Dim X As Integer
If Me.HfgEfectivo.Rows < 1 Then
    X = 0
End If
End Sub

Private Sub HfgEfectivo_DblClick()
If Me.HfgEfectivo.Rows > 0 Then
    Procedencia = buscar
    frmdetalle.Show
End If
End Sub

Private Sub impresion_reporte()
    Dim tttarjeta As Double
    Call CargaDefConfigEpsonTM
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Font.name = "FontB11"Draft 17cpi
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    '
        
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); KEY_DIRECCION
    Printer.Print Tab(1); "RUC :" & KEY_RUC
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); "ARQUEO DE CAJA"
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); "FECHA:" & KEY_FECHA
    If Me.ChkAlmacen.Value = 1 Then
        Printer.Print Tab(1); "TIENDA :" + Me.DtcAlmacen.Text
    Else
        Printer.Print Tab(1); "ACUMULADO TOTAL TODAS LAS TIENDAS"
    End If
    If Me.chkOperador.Value = 1 Then
        Printer.Print Tab(1); "CAJERA :" + Me.DtcOperador.Text
           
    End If
    tTotal = Val(Me.lblTotalEfectivo.Caption)
    tttarjeta = Val(Me.lblTotalMastercard.Caption) + Val(Me.lblTotalVisa.Caption)
    
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); Mid("ANULADAS" + Space(20), 1, 12) & ":" & Mid(Format(Val(Me.lblNumanuladas.Caption), "#,##0.00") + Space(10), 1, 5) + Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalAnuladas.Caption), "#,##0.00")
    Printer.Print Tab(1); Mid("EFECTIVO" + Space(20), 1, 12) & ":" & Mid(Format(Val(Me.lblNumEfectivo.Caption), "#,##0.00") + Space(10), 1, 5) + Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalEfectivo.Caption), "#,##0.00")
    Printer.Print Tab(1); Mid("VISA" + Space(20), 1, 12) & ":" & Mid(Format(Val(Me.lblNumVisa.Caption), "#,##0.00") + Space(10), 1, 5) + Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalVisa.Caption), "#,##0.00")
    Printer.Print Tab(1); Mid("MASTERCARD" + Space(20), 1, 12) & ":" & Mid(Format(Val(Me.LblnumMastercard.Caption), "#,##0.00") + Space(10), 1, 5) + Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalMastercard.Caption), "#,##0.00")
    Printer.Print Tab(1); Mid("CREDITOS" + Space(20), 1, 12) & ":" & Mid(Format(Val(Me.lblnumCreditos.Caption), "#,##0.00") + Space(10), 1, 5) + Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalCreditos.Caption), "#,##0.00")
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); Mid("TOTAL EFECTIVO" + Space(20), 1, 16) & ":" & Space(1) + "======>" + Space(1) + Format(tTotal, "#,##0.00")
    Printer.Print Tab(1); Mid("TOTAL TARJETAS" + Space(20), 1, 16) & ":" & Space(1) + "======>" + Space(1) + Format(tttarjeta, "#,##0.00")
    Printer.Print Tab(1); Mid("TOTAL CREDITOS" + Space(20), 1, 16) & ":" & Space(1) + "======>" + Space(1) + Format(Val(Me.lblTotalCreditos.Caption), "#,##0.00")
    Printer.Print ""
    Printer.Print Tab(1); Mid("ACUMULADO CAJA" + Space(20), 1, 16) & ":" & Space(1) + "======>" + Space(1) + Format(tTotal + Val(Me.lblTotalCreditos.Caption), "#,##0.00")
    Printer.Print Tab(1); Mid("ACUMULADO TOTAL" + Space(20), 1, 16) & ":" & Space(1) + "======>" + Space(1) + Format(tttarjeta + tTotal + Val(Me.lblTotalCreditos.Caption), "#,##0.00")
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); UCase(EnLetras(tTotal + tttarjeta + Val(Me.lblTotalCreditos.Caption)))
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(10); "----------------------------------"
    Printer.Print Tab(10); Mid(Space(10) + "ENCARGADO" + Space(10), 1, 20)
    Printer.Print Tab(10); "(" + Space(5) + Mid(KEY_VENDEDOR + Space(10), 1, 30) + ")"
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(1); Mid("HORA PROCESO" + Space(10), 1, 12) & ":" & str(Time)
    Printer.EndDoc
    
    Exit Sub
End Sub








Public Sub llenar_anuladas(ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
Dim ttanuladas As Double
Dim turno As String
turno = ""
If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If
ttanuladas = 0
If id_alm = "00001" Then
    strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' AND id_vendedor LIKE '%" & id_usuario & "%'  AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' AND anulado='si' AND turno LIKE '%" & Trim(turno) & "%'"
Else
    strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' AND id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' AND anulado='si' AND turno LIKE '%" & Trim(turno) & "%'"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    Me.lblTotalAnuladas.Caption = 0
    For i = 0 To rst.RecordCount - 1
        ttanuladas = rst("total") + ttanuladas
        rst.MoveNext
    Next i
End If
Me.lblTotalAnuladas.Caption = Format(ttanuladas, "###0.00")
Me.lblNumanuladas.Caption = str(rst.RecordCount)
Set rst = Nothing
End Sub
Public Sub Llenar_mastercard(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir
Dim tTotal As Double
Dim turno As String
  turno = ""
'If Me.OptManana.Value = True Then
    turno = "M"
'End If
'If Me.OptTarde.Value = True Then
    turno = "T"
'End If
'strCadena = "SELECT V.id_venta,V.fecha_emision,CONCAT(C.doc_abrev,':',V.serie,'-',V.numero)as comprobante,P.nombre_completo,M.monto,V.anulado " & _
"FROM movimiento_venta V,movimiento_venta_monto M, comprobantes C,persona P  WHERE V.id_venta=M.id_venta AND V.id_doc=C.id_doc AND V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
" AND V.id_cliente=P.dni AND V.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND M.id_tarjeta='02' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and V.id_vendedor LIKE '%" & id_usuario & "%' AND V.id_alm LIKE '%" & id_alm & "%' AND  V.turno LIKE '%" & Trim(turno) & "%' ORDER BY V.serie,V.numero ASC"
strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_ultimate WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_tarjeta='02' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
    

Call ConfiguraRst(strCadena)
Me.LblnumMastercard.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Me.LblnumMastercard.Caption = "0"
    Me.lblTotalVisa.Caption = "0.00"
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1000
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2400
            Grilla.ColWidth(4) = 1100
            Grilla.ColWidth(5) = 1900
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & Format(rst("monto_caja"), "###0.00") & vbTab & rst("vendedor")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
               Me.lblNumanuladas.Caption = Val(Me.lblNumanuladas.Caption) + 1
                    For k = 0 To 4
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                tTotal = tTotal + rst("monto_caja")
            End If
            
           
            rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      Me.lblTotalMastercard.Caption = Format(tTotal, "###0.000")
      For k = 3 To 4
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing







End Sub
Public Sub Llenar_pago_credito(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir
Dim tTotal As Double

'strCadena = "SELECT V.id_venta,V.fecha_emision,CONCAT(C.doc_abrev,':',V.serie,'-',V.numero)as comprobante,P.nombre_completo,M.monto,V.anulado " & _
"FROM movimiento_venta V,movimiento_venta_monto M, comprobantes C,persona P  WHERE V.id_venta=M.id_venta AND V.id_doc=C.id_doc AND V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
" AND V.id_cliente=P.dni AND V.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND M.id_tarjeta='02' AND V.id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  V.id_vendedor LIKE '%" & id_usuario & "%' AND V.id_alm LIKE '%" & id_alm & "%' ORDER BY V.serie,V.numero ASC"
Exit Sub
strCadena = "SELECT * FROM mis_cuentas_det M,persona P WHERE M.id_persona=P.dni AND  M.dni_save LIKE '%" & id_usuario & "%' AND ruc='" & KEY_RUC & "' AND id_alm LIKE '%" & id_alm & "%' AND fecha>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' "
Call ConfiguraRst(strCadena)
Me.lblnumCreditos.Caption = str(rst.RecordCount)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.lblTotalCreditos.Caption = Format(0, "#,##0.00")
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1100
            Grilla.ColWidth(2) = 1800
            Grilla.ColWidth(3) = 2500
            Grilla.ColWidth(4) = 1300
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_movimiento") & vbTab & rst("fecha") & vbTab & Mid(rst("glosa"), 7, 19) & vbTab & rst("nombre_completo") & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
                    For k = 0 To 4
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                tTotal = tTotal + rst("monto")
            End If
            
           
            rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      Me.lblTotalCreditos.Caption = Format(tTotal, "###0.000")
      For k = 0 To 4
            Grilla.col = 4
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub Llenar_visa(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir
Dim tTotal As Double
Dim turno As String
  turno = ""
'If Me.OptManana.Value = True Then
   ' turno = "M"
'End If
'If Me.OptTarde.Value = True Then
   ' turno = "T"
'End If
'strCadena = "SELECT V.id_venta,V.fecha_emision,CONCAT(C.doc_abrev,':',V.serie,'-',V.numero)as comprobante,P.nombre_completo,M.monto,V.anulado " & _
"FROM movimiento_venta V,movimiento_venta_monto M, comprobantes C,persona P  WHERE V.id_venta=M.id_venta AND V.id_doc=C.id_doc AND V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
" AND V.id_cliente=P.dni AND V.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND M.id_tarjeta='01' and  V.id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%'  AND V.id_vendedor LIKE '%" & id_usuario & "%' AND V.id_alm LIKE '%" & id_alm & "%' AND  V.turno LIKE '%" & Trim(turno) & "%' ORDER BY V.serie,V.numero ASC"

    strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_ultimate WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_tarjeta='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
    
    
Call ConfiguraRst(strCadena)
Me.lblNumVisa.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblTotalVisa.Caption = Format(0, "#,##0.00")
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1000
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2400
            Grilla.ColWidth(4) = 1100
            Grilla.ColWidth(5) = 1900
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & Format(rst("monto_caja"), "###0.00") & vbTab & rst("vendedor")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
                    For k = 0 To 4
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                tTotal = tTotal + rst("monto_caja")
            End If
            
           
            rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      Me.lblTotalVisa.Caption = Format(tTotal, "###0.000")
      For k = 0 To 4
            Grilla.col = 4
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Public Sub llenar_devoluciones(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir
Dim tTotal As Double
Dim turno As String
turno = ""
If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If

If Me.DtcAlmacen.BoundText = "00001" Then
    strCadena = "SELECT V.id_venta,V.fecha_emision,documento as comprobante,ncliente as nombre_completo,V.total ,V.anulado,P.nombre_completo as digitador " & _
    "FROM movimiento_venta V,persona P WHERE V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND V.ruc='" & KEY_RUC & "' AND V.id_vendedor=P.dni AND V.id_forma_pago='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and V.id_vendedor LIKE '%" & id_usuario & "%' AND V.turno LIKE '%" & Trim(turno) & "%' AND id_doc='0205' ORDER BY documento ASC"
Else
    strCadena = "SELECT V.id_venta,V.fecha_emision,documento as comprobante,ncliente as nombre_completo,V.total ,V.anulado,P.nombre_completo as digitador " & _
    "FROM movimiento_venta V,persona P WHERE V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND V.ruc='" & KEY_RUC & "' AND V.id_vendedor=P.dni AND V.id_forma_pago='01' AND V.id_vendedor LIKE '%" & id_usuario & "%' AND V.id_alm LIKE '%" & id_alm & "%' AND V.turno LIKE '%" & Trim(turno) & "%' AND  id_doc='0205' ORDER BY documento ASC"
End If

Call ConfiguraRst(strCadena)
Me.lblnumdevoluciones.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblnumdevoluciones.Caption = "0"
    Me.lblTotalDevoluciones.Caption = "0.00"
    Exit Sub
End If
   Grilla.Rows = 1
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1100
            Grilla.ColWidth(2) = 1900
            Grilla.ColWidth(3) = 3100
            Grilla.ColWidth(4) = 1000
            Grilla.ColWidth(5) = 1800
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & Format(rst("total"), "###0.00") & vbTab & rst("digitador")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
               Me.lblNumanuladas.Caption = Val(Me.lblNumanuladas.Caption) + 1
                    For k = 0 To 5
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                    tTotal = tTotal + rst("total")
            End If
            
           
            rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
 Me.lblTotalDevoluciones.Caption = Format(tTotal, "###0.000")
      For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"

End Sub
Public Sub llenar_garantias(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String)
On Error GoTo salir
Dim tTotal As Double
Dim turno As String
turno = ""
If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If

If Me.DtcAlmacen.BoundText = "00001" Then
    strCadena = "SELECT V.id_venta,V.fecha_emision,documento as comprobante,ncliente as nombre_completo,V.total ,V.anulado,P.nombre_completo as digitador " & _
    " FROM movimiento_venta V,movimiento_venta_detalle D,persona P WHERE V.id_vendedor= P.dni AND  V.id_venta=D.id_movimiento AND D.id_producto='99527' AND V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND V.ruc='" & KEY_RUC & "' AND V.id_vendedor=P.dni AND V.id_forma_pago='01' AND V.id_vendedor LIKE '%" & id_usuario & "%' AND V.turno LIKE '%" & Trim(turno) & "%' AND id_doc='0205' ORDER BY documento ASC"
Else
    strCadena = "SELECT V.id_venta,V.fecha_emision,documento as comprobante,ncliente as nombre_completo,V.total ,V.anulado,P.nombre_completo as digitador " & _
    "FROM movimiento_venta V,movimiento_venta_detalle D,persona P WHERE V.id_vendedor= P.dni AND V.id_venta=D.id_movimiento AND D.id_producto='99527' V.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND V.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND V.ruc='" & KEY_RUC & "' AND V.id_vendedor=P.dni AND V.id_forma_pago='01' AND V.id_vendedor LIKE '%" & id_usuario & "%' AND V.id_alm LIKE '%" & id_alm & "%' AND V.turno LIKE '%" & Trim(turno) & "%' AND  id_doc='0205' ORDER BY documento ASC"
End If

Call ConfiguraRst(strCadena)
Me.lblnumdevoluciones.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblnumdevoluciones.Caption = "0"
    Me.lblTotalDevoluciones.Caption = "0.00"
    Exit Sub
End If
   Grilla.Rows = 1
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1100
            Grilla.ColWidth(2) = 1900
            Grilla.ColWidth(3) = 3100
            Grilla.ColWidth(4) = 1000
            Grilla.ColWidth(5) = 1800
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & Format(rst("total"), "###0.00") & vbTab & rst("digitador")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
               Me.lblNumanuladas.Caption = Val(Me.lblNumanuladas.Caption) + 1
                    For k = 0 To 5
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                    tTotal = tTotal + rst("total")
            End If
            
           
            rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
 Me.lblTotalDevoluciones.Caption = Format(tTotal, "###0.000")
      For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"

End Sub

Public Sub Llenar_efectivo(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir

Dim tTotal As Double
Dim turno As String
turno = ""
If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If

  '  strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_v2 WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_detalle='01' and id='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
    
    'strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,(monto_caja-if(id_doc='0007',function_get_nota_credito_pago(id_venta,ruc)*-1,function_get_nota_credito(id_venta)) ) as monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_ultimate WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_detalle='01' and id='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
   
 strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_ultimate WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_detalle='01' and id='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
   
   
Call ConfiguraRst(strCadena)
Me.lblNumEfectivo.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblNumEfectivo.Caption = "0"
    Me.lblTotalEfectivo.Caption = "0.00"
    Exit Sub
End If
   
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1000
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2400
            Grilla.ColWidth(4) = 1100
            Grilla.ColWidth(5) = 1900
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & Format(rst("monto_caja"), "###0.00") & vbTab & rst("vendedor")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
               Me.lblNumanuladas.Caption = Val(Me.lblNumanuladas.Caption) + 1
                    For k = 0 To 5
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                    tTotal = tTotal + rst("monto_caja")
            End If
        rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      Me.lblTotalEfectivo.Caption = Format(tTotal, "###0.000")
      txttotalsistema.Text = Format(tTotal, "###0.000")
      For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub llenar_deposito_cuenta(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
On Error GoTo salir

Dim tTotal As Double
Dim turno As String
turno = ""
If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If

  '  strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_v2 WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_detalle='01' and id='01' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
    
    strCadena = "SELECT id_venta,fecha_emision,documento ,ncliente ,monto_caja,total,anulado,vendedor " & _
    "FROM view_reporte_detallado_ultimate WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
    " AND ruc='" & KEY_RUC & "' AND id_detalle='09' AND id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and id_vendedor LIKE '%" & id_usuario & "%' AND id_alm LIKE '%" & id_alm & "%' AND turno LIKE '%" & Trim(turno) & "%' AND afecta_caja='si' and afecta_factura='no' and id_recibo='0' ORDER BY documento ASC"
   '
Call ConfiguraRst(strCadena)
Me.lblNumEfectivo.Caption = str(rst.RecordCount)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblNumEfectivo.Caption = "0"
    Me.lblTotalEfectivo.Caption = "0.00"
    Exit Sub
End If
   
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1000
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2400
            Grilla.ColWidth(4) = 1100
            Grilla.ColWidth(5) = 1900
        Next
        cabecera = "ID_VENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PACIENTE" & vbTab & "MONTO" & vbTab & "DIGITADOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & Format(rst("monto_caja"), "###0.00") & vbTab & rst("vendedor")
            Grilla.AddItem Fila
            Fila = ""
            If rst("anulado") = "si" Then
               Me.lblNumanuladas.Caption = Val(Me.lblNumanuladas.Caption) + 1
                    For k = 0 To 5
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
            Else
                    tTotal = tTotal + rst("monto_caja")
            End If
        rst.MoveNext
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      Me.lblTotalEfectivo.Caption = Format(tTotal, "###0.000")
      For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
     
     
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub Llenar_tarde(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String)
'On Error GoTo SALIR
Dim tTotal As Double
Dim Tpacientes As Double
Dim turno As String
Dim Acumulado As Double
  turno = ""
'If Me.OptManana.Value = True Then
    turno = "M"
'End If
'If Me.OptTarde.Value = True Then
    turno = "T"
'End If
'Strcadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' AND E.id_sucursal='" & KEY_ALM & "'"
If KEY_ALM = "00001" Then
    strCadena = "SELECT DISTINCT V.id_vendedor as Codigo,P.nombre_completo as Descripcion  FROM movimiento_venta V,persona P WHERE V.id_vendedor=P.dni AND V.ruc='" & KEY_RUC & "'  AND V.fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "'"
Else
    strCadena = "SELECT DISTINCT V.id_vendedor as Codigo,P.nombre_completo as Descripcion  FROM movimiento_venta V,persona P WHERE V.id_vendedor=P.dni AND V.ruc='" & KEY_RUC & "' AND V.id_alm='" & KEY_ALM & "' AND V.fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblNumEfectivo.Caption = "0"
    Me.lblTotalEfectivo.Caption = "0.00"
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 1200
            Grilla.ColWidth(1) = 3500
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2000
        Next
        cabecera = "DNI" & vbTab & "DESCRIPCION" & vbTab & "Nº PACIENTES" & vbTab & "MONTO ACUMULADO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        Chart.TitleText = "ESTADISTICA DE PRODUCCION"
        Me.Chart.RowCount = rst.RecordCount
        With Chart.DataGrid
        rst.MoveFirst
        
          For i = 0 To rst.RecordCount - 1
            If KEY_ALM = "00001" Then
                strCadena = "SELECT sum(total),count(*) FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
                " AND ruc='" & KEY_RUC & "' AND id_forma_pago='01' AND id_vendedor LIKE '%" & rst("Codigo") & "%'   AND anulado='no'"
            Else
                strCadena = "SELECT sum(total),count(*) FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' " & _
                " AND ruc='" & KEY_RUC & "' AND id_forma_pago='01' AND id_vendedor LIKE '%" & rst("Codigo") & "%' AND id_alm LIKE '%" & id_alm & "%'  AND anulado='no'"
            End If
            Call ConfiguraRstT(strCadena)
            If IsNull(rstT(0)) = True Then
                tTotal = 0
            Else
                tTotal = rstT(0)
            End If
            Tpacientes = Tpacientes + rstT(1)
           .RowLabel(i + 1, 1) = rst("Descripcion")
           .SetSize rst.RecordCount, 1, rst.RecordCount, 1
            Acumulado = Acumulado + tTotal
            Fila = rst("Codigo") & vbTab & rst("Descripcion") & vbTab & rstT(1) & vbTab & Format(tTotal, "###0.00")
            Grilla.AddItem Fila
            Fila = ""
                      
           
            rst.MoveNext
        Next i
        End With
          Fila = "" & vbTab & "RECAUDACION ACUMULADA" & vbTab & Tpacientes & vbTab & Format(Acumulado, "###0.00")
          Grilla.AddItem Fila
            '&HC0FFFF
            For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
        Next k
     
'Exit Sub
'SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Sub llenar_imc_historial(ByVal dni As String)
'Height = 5640: Width = 7575
Chart.Visible = True

'If Len(cboYear.text) = 4 Then
    ' Filtramos los Registros de la BD en base a el año seleccionado
    strCadena = "SELECT fecha,peso FROM persona_peso WHERE dni='" & dni & "'"
    Call ConfiguraRstT(strCadena)
    Chart.TitleText = "VARIACION DE PESO"
    Me.Chart.RowCount = rstT.RecordCount
    With Chart.DataGrid
        For i = 1 To rstT.RecordCount
        .RowLabel(i, 1) = rstT("fecha")
        .SetSize rstT.RecordCount, 1, rstT.RecordCount, 1
        rstT.MoveNext
    Next i
    End With

End Sub
Public Sub Llenar_reporte(ByVal Grilla As MSHFlexGrid, ByVal id_usuario As String, ByVal id_alm As String, ByVal in_ventanilla As String)
'On Error GoTo SALIR
Dim tTotal As Double
Dim Tpacientes As Double
Dim turno As String
Dim Acumulado As Double
Dim doc_ini As String
Dim doc_fin As String



turno = ""

If Me.chkturno.Value = 1 Then
     turno = Replace(Me.DtcTurno.BoundText, "'", "''")
End If

'If Me.chkAlmacen.Value = 1 Then
 '    id_alm = Replace(Me.DtcAlmacen.BoundText, "'", "''")
'End If
'If Me.chkOperador.Value = 1 Then
 '   id_usuario = Replace(Me.DtcOperador.BoundText, "'", "''")
'End If


If id_alm = "00001" Then
    strCadena = "SELECT DISTINCT V.id_alm as Codigo,A.descripcion as Descripcion  FROM movimiento_venta V,almacen A WHERE V.id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  V.id_alm=A.id_alm AND V.ruc='" & KEY_RUC & "'  AND V.fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND  V.fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND A.ruc='" & KEY_RUC & "' AND turno LIKE '%" & turno & "%' AND V.id_forma_pago='01' AND V.id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no' AND (id_doc='0001' OR id_doc='0003') ORDER BY A.id_alm"
Else
    strCadena = "SELECT DISTINCT V.id_alm as Codigo,A.descripcion as Descripcion  FROM movimiento_venta V,almacen A WHERE V.id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and V.id_alm LIKE '%" & id_alm & "%' AND V.id_alm=A.id_alm AND V.ruc='" & KEY_RUC & "' AND A.ruc='" & KEY_RUC & "'  AND V.fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND V.fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND V.id_forma_pago='01' AND anulado='no' AND V.id_vendedor LIKE '%" & id_usuario & "%' AND (id_doc='0001' OR id_doc='0003') ORDER BY A.id_alm"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    'Me.lblNumEfectivo.Caption = "0"
    'Me.lblTotalEfectivo.Caption = "0.00"
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 3500
            Grilla.ColWidth(1) = 3000
            Grilla.ColWidth(2) = 1200
            Grilla.ColWidth(3) = 1500
        Next
        cabecera = "VENTANILLA" & vbTab & "COMPROBANTES" & vbTab & "Nº PACIENTES" & vbTab & "MONTO ACUMULADO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        Chart.TitleText = "ESTADISTICA DE PRODUCCION"
        Me.Chart.RowCount = rst.RecordCount
        
        With Chart.DataGrid
        
        rst.MoveFirst
        
          For i = 0 To rst.RecordCount - 1
            If KEY_ALM = "00001" Then
              strCadena = "SELECT CONCAT(serie,'-',numero) FROM movimiento_venta WHERE  id_alm='" & rst("Codigo") & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND  fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no' AND (id_doc='0001' OR id_doc='0003')  ORDER BY serie ASC,numero ASC LIMIT 0,1"
              Call ConfiguraRstL(strCadena)
              If rstL.RecordCount > 0 Then
                 boleta_ini = "[" & Space(1) & rstL(0) & Space(1) & "]"
              End If
              strCadena = "SELECT CONCAT(serie,'-',numero) FROM movimiento_venta WHERE  id_alm='" & rst("Codigo") & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no' AND (id_doc='0001' OR id_doc='0003') ORDER BY serie DESC,numero DESC LIMIT 0,1"
              Call ConfiguraRstL(strCadena)
              If rstL.RecordCount > 0 Then
                 boleta_fin = "[" & Space(1) & rstL(0) & Space(1) & "]"
              End If
              
              strCadena = "SELECT sum(total),count(*) FROM movimiento_venta WHERE  id_alm='" & rst("Codigo") & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no'  AND (id_doc='0001' OR id_doc='0003') ORDER BY serie ASC,numero ASC LIMIT 0,1"
                
            Else
                 strCadena = "SELECT CONCAT(serie,'-',numero) FROM movimiento_venta WHERE  id_alm='" & id_alm & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%'  AND anulado='no' AND (id_doc='0001' OR id_doc='0003') ORDER BY serie ASC,numero ASC LIMIT 0,1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    boleta_ini = "[" & Space(1) & rstL(0) & Space(1) & "]"
                End If
                strCadena = "SELECT CONCAT(serie,'-',numero) FROM movimiento_venta WHERE  id_alm='" & id_alm & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no' AND (id_doc='0001' OR id_doc='0003') ORDER BY serie DESC,numero DESC LIMIT 0,1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    boleta_fin = "[" & Space(1) & rstL(0) & Space(1) & "]"
                End If
              
              strCadena = "SELECT sum(total),count(*) FROM movimiento_venta WHERE  id_alm='" & id_alm & "' AND ruc='" & KEY_RUC & "'  AND fecha_emision>='" & formato_fecha(Me.DtpDesde.Value) & "' AND fecha_emision<='" & formato_fecha(Me.DtpHasta.Value) & "' AND turno LIKE '%" & turno & "%' AND id_forma_pago='01' AND id_vendedor LIKE '%" & id_usuario & "%' AND anulado='no' AND (id_doc='0001' OR id_doc='0003') ORDER BY serie ASC,numero ASC LIMIT 0,1"
            End If
            Call ConfiguraRstT(strCadena)
            If IsNull(rstT(0)) = True Then
                tTotal = 0
            Else
                tTotal = rstT(0)
            End If
            Tpacientes = Tpacientes + rstT(1)
           .RowLabel(i + 1, 1) = rst("Descripcion")
           .SetSize rst.RecordCount, 1, rst.RecordCount, 1
           '.SetData 1, 1, rstT(1), 0
            Acumulado = Acumulado + tTotal
            Fila = rst("descripcion") & vbTab & boleta_ini & Space(3) & boleta_fin & vbTab & rstT(1) & vbTab & Format(tTotal, "###0.00")
            Grilla.AddItem Fila
            Fila = ""
                      
           
            rst.MoveNext
        Next i
        End With
        
          Fila = "" & vbTab & "RECAUDACION ACUMULADA" & vbTab & Tpacientes & vbTab & Format(Acumulado, "###0.00")
          Grilla.AddItem Fila
            '&HC0FFFF
            For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
        Next k
     
'Exit Sub
'SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub LlenarComprobantes(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim tTotal As Double
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2300
           Grilla.ColWidth(3) = 0
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1500
         Next
        cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "DOCUMENTO" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "VENDEDOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("idVenta") & vbTab & rst("Fecha") & vbTab & rst("Numero") & vbTab & rst("Persona") & vbTab & Format(rst("nTotalVenta"), "#,##0.00") & vbTab & rst("Usuario")
            Grilla.AddItem Fila
            
            tTotal = tTotal + rst("nTotalVenta")
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 5 To 5
            Grilla.col = 4
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
'salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub LlenarAnulada(ByVal HfdGrillaCont As MSHFlexGrid)
        Call ConfiguraRst(strCadena)
        HfdGrillaCont.Rows = rst.RecordCount
        
        Set HfdGrillaCont.Recordset = rst
       
            
        
        HfdGrillaCont.ColWidth(0) = 1000
        HfdGrillaCont.ColWidth(1) = 2100
        HfdGrillaCont.ColWidth(2) = 4200
        HfdGrillaCont.ColWidth(3) = 1500
        Call DarFormatoFecha(HfdGrillaCont, 0)
        Call DarFormato(HfdGrillaCont, 3)
        Set HfdGrillaCont = Nothing
        Set rst = Nothing
                
End Sub
Public Sub SumarTotal(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
Dim i As Integer
Dim Total As Double
  Total = 0
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  For i = 0 To longitud - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.00")
  Next i
  HfdPrecio.Refresh
End Sub
Private Sub Impresion_contable(ByVal FechaIni As Date, ByVal FechaFin As Date)
Dim nanuladas As Integer
Dim nefectivo As Integer
Dim ncredito As Integer
Dim ndebito As Integer
Dim nTotal As Integer
Dim ndeudas As Integer
Dim tanuladas As Double
Dim tefectivo As Double
Dim tcredito As Double
Dim tdebito As Double
Dim tdeudas As Double

Dim nUsuario As String
'Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
 '   Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Font.name = "Draft 17cpi"
  '  Printer.Font.name = "Times New Roman"
   ' Printer.Font.Size = 12
    Anulado = "F"
   StrAlmacen = ""
        StrOperador = ""
        If Me.ChkAlmacen.Value = 1 Then
        StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
        End If
        
        If Me.chkOperador.Value = 1 Then
         StrOperador = Replace(Me.DtcOperador.BoundText, "'", "''")
        End If
    
        strCadena = "SELECT SUM(nTotalVenta) ,COUNT(*)  FROM DocumentoVenta WHERE idFormaPago='0001' AND dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND id_usuario LIKE '" & StrOperador & "%' AND Anulado='F' AND estado='Cancelado' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND DocumentoVenta.doc_cod<>'0010' AND DocumentoVenta.doc_cod<>'0097'"
        Call ConfiguraRst(strCadena)
        
        If rst(1) > 0 Then
            tefectivo = rst(0)
            nefectivo = rst(1)
        Else
            tefectivo = 0
            nefectivo = 0
        End If
        Set rst = Nothing
          
    strCadena = "SELECT SUM(nTotalVenta) ,COUNT(*)  FROM DocumentoVenta WHERE  dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND id_usuario LIKE '" & StrOperador & "%' AND Anulado='V' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
        Call ConfiguraRst(strCadena)
         If rst(1) > 0 Then
           tanuladas = rst(0)
           nanuladas = rst(1)
        Else
          tanuladas = 0
          nanuladas = 0
        End If
        Set rst = Nothing
        strCadena = "SELECT SUM(nTotalVenta) ,COUNT(*)  FROM DocumentoVenta WHERE idFormaPago='0002' AND dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND id_usuario LIKE '" & StrOperador & "%' AND Anulado='F' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
        Call ConfiguraRst(strCadena)
          If rst(1) > 0 Then
            tcredito = rst(0)
            ncredito = rst(1)
        Else
            tcredito = 0
            ncredito = 0
        End If
        Set rst = Nothing
        strCadena = "SELECT SUM(nTotalVenta) ,COUNT(*)  FROM DocumentoVenta WHERE idFormaPago='0003' AND dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND id_usuario LIKE '" & StrOperador & "%' AND Anulado='F' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
        Call ConfiguraRst(strCadena)
         If rst(1) > 0 Then
             tdebito = rst(0)
             ndebito = rst(1)
        Else
            tdebito = 0
            tdebito = 0
        End If
    Set rst = Nothing
         strCadena = "SELECT SUM(nTotalVenta) ,COUNT(*)  FROM DocumentoVenta WHERE idFormaPago='0001' AND dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND " & _
        "dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND id_usuario LIKE '" & StrOperador & "%' AND Anulado='F' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND DocumentoVenta.doc_cod='0010'"
        Call ConfiguraRst(strCadena)
        
        If rst(1) > 0 Then
            tdeudas = rst(0)
            ndeudas = rst(1)
        Else
           tdeudas = 0
            ndeudas = 0
        End If
        Set rst = Nothing
        
    Dim cod_reporte As String
    
    Dim tTotal As Double
    Dim nombre_user As String
    nTotal = ndeudas + nefectivo + ncredito + ndeudas + ndebito
    tTotal = tdeudas + tefectivo + tcredito + tdebito
    strCadena = "SELECT * FROM Seguridad WHERE IdUsuario ='" & Trim(KEY_USUARIO) & "' "
    Call ConfiguraRst(strCadena)
    nombre_user = rst("Nombre")
    Set rst = Nothing
    Dim operadortext As String
    If Me.chkOperador.Value = 1 Then
        operadortext = Me.DtcOperador.Text
    Else
       operadortext = "REPORTE GLOBAL"
    End If
    strCadena = "SELECT * FROM CobranzaDiaria WHERE id_usuario='" & Trim(KEY_USUARIO) & "' ORDER BY id_reporte DESC"
    Call ConfiguraRst(strCadena)
    cod_reporte = GeneraCodigo(4)
    strCadena = "INSERT INTO CobranzaDiaria VALUES('" & Trim(cod_reporte) & "','" & Me.DtpDesde.Value & "','" & Me.DtpHasta.Value & "','" & nanuladas & "','" & nefectivo & "','" & ncredito & "'," & _
    "'" & ndebito & "','" & ndeudas & "','" & tanuladas & "','" & tefectivo & "','" & tcredito & "','" & tdebito & "','" & tdeudas & "','" & nTotal & "','" & tTotal & "','" & Trim(nombre_user) & "','" & Trim(operadortext) & "','" & tefectivo + tdeudas & "','" & KEY_USUARIO & "')"
    EjecutaRST (strCadena)
    Set RstEjecuta = Nothing
    strCadena = "SELECT * FROM CobranzaDiaria WHERE id_reporte='" & Trim(cod_reporte) & "' AND id_usuario='" & KEY_USUARIO & "'"
    Call ConfiguraRst(strCadena)
    Set DtrCobranza.DataSource = rst
    DtrCobranza.Show
    Set rst = Nothing
  
  
End Sub







Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TxtBusquedaRapida_Change()
strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' AND E.id_medico='no' AND E.id_cargo<>'00025' AND E.id_cargo<>'00015' AND P.nombre_completo LIKE '%" & Trim(Me.TxtBusquedaRapida.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcOperador)
End Sub


Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "UPDATE arqueo_detalle SET cantidad='" & Val(Me.txtcantidad.Text) & "' WHERE id='" & Val(Me.HfArqueo.TextMatrix(Me.HfArqueo.Row, 0)) & "'"
    Call Execute_Sql(strCadena)
    Call Me.llenar_arqueo(Me.HfArqueo, Val(Me.txt_idarqueo.Text))
    Me.frmcantidad.Visible = False
End If
End Sub

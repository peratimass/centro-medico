VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteProducto 
   BorderStyle     =   0  'None
   Caption         =   "Reporte de Productos"
   ClientHeight    =   9135
   ClientLeft      =   540
   ClientTop       =   315
   ClientWidth     =   16695
   Icon            =   "FrmReporteProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8840
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "     REPORTES GENERALES"
      TabPicture(0)   =   "FrmReporteProducto.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "prg_reporte"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdUpdateCosto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDetallado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd_cerrar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "opt_listado_general"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "opt_producto_mas_vendido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "opt_producto_mayor_utilidad"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "opt_producto_menor_utilidad"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "opt_clientes_mas_ventas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "opt_clientes_menos_ventas"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "opt_cliente_mayor_utilidad"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "opt_cliente_menor_utilidad"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmd_reporte"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "opt_producto_menos_vendidos"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "opt_ventas_vendedor"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "opt_venta_categoria"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Opt_reporte102"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Opt_reporte103"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Opt_reporte104"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Opt_reporte105"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Opt_reporte106"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Opt_reporte107"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Opt_reporte108"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Opt_reporte109"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Opt_reporte1011"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Opt_reporte1012"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Opt_reporte1013"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "opt_reporte101"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "opt_series_disponibles"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "opt_series_vendidas"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "opt_kardex_linea"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "opt_resumen"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "opt_obsequios"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "chk_producto_stock"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "chk_todos"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "opt_catalogo"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "  REPORTES COMPARATIVOS"
      TabPicture(1)   =   "FrmReporteProducto.frx":08A4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Opt_comparativo_diario"
      Tab(1).Control(2)=   "Opt_comparativo_mensual"
      Tab(1).Control(3)=   "Opt_comparativo_mensual_cantidad"
      Tab(1).Control(4)=   "Opt_comparativo_diario_cantidad"
      Tab(1).Control(5)=   "cmdReporte_comparativo"
      Tab(1).Control(6)=   "cmdcerrar_comparativo"
      Tab(1).Control(7)=   "Shape2"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "RENDIMIENTO"
      TabPicture(2)   =   "FrmReporteProducto.frx":0E3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "opt_produccion"
      Tab(2).Control(1)=   "opt_gastos_personal"
      Tab(2).Control(2)=   "opt_gasto_personal"
      Tab(2).Control(3)=   "opt_planilla_cobrador"
      Tab(2).Control(4)=   "Frame3"
      Tab(2).Control(5)=   "OptReporteventas"
      Tab(2).Control(6)=   "cmdreorterendimiento"
      Tab(2).Control(7)=   "cmdCerrarRendimiento"
      Tab(2).Control(8)=   "prg_avance"
      Tab(2).Control(9)=   "Shape4"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "GERENCIALES"
      TabPicture(3)   =   "FrmReporteProducto.frx":0E5A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(2)=   "OptReporteMensual"
      Tab(3).Control(3)=   "OptResumen1"
      Tab(3).Control(4)=   "cmdmensual"
      Tab(3).Control(5)=   "Shape5"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "LIQUIDACION/LOGISTICA"
      TabPicture(4)   =   "FrmReporteProducto.frx":0E76
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame8"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).Control(2)=   "framemayor"
      Tab(4).Control(3)=   "Frame6"
      Tab(4).Control(4)=   "Frame5"
      Tab(4).Control(5)=   "Shape6"
      Tab(4).ControlCount=   6
      Begin VB.OptionButton opt_catalogo 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CATALOGO DE PRECIOS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   159
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chk_todos 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "TODOS"
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
         Height          =   320
         Left            =   4200
         TabIndex        =   152
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chk_producto_stock 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "CON STOCK"
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
         Height          =   320
         Left            =   3000
         TabIndex        =   151
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton opt_produccion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "REPORTE PRODUCCION PERSONAL"
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
         Height          =   320
         Left            =   -74520
         TabIndex        =   148
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ESTADO DE CUENTA GENERAL"
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
         Height          =   1815
         Left            =   -66000
         TabIndex        =   145
         Top             =   1080
         Width           =   6495
         Begin VB.CheckBox chk_personal 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "PERSONAL"
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
            Height          =   255
            Left            =   1440
            TabIndex        =   147
            Top             =   480
            Width           =   1215
         End
         Begin VitekeySoft.ChameleonBtn cmdEstadoCuentaGeneral 
            Height          =   495
            Left            =   1440
            TabIndex        =   146
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "ESTADO DE CUENTA GENERAL"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":0E92
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
      Begin VB.OptionButton opt_gastos_personal 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMPRAS Y GASTOS POR PERSONAL"
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
         Height          =   320
         Left            =   -74520
         TabIndex        =   144
         Top             =   3765
         Width           =   4095
      End
      Begin VB.OptionButton opt_obsequios 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "OBSEQUIOS DESCUENTOS Y BONIFICACIONES"
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
         Height          =   320
         Left            =   7440
         TabIndex        =   138
         Top             =   2430
         Width           =   4815
      End
      Begin VB.OptionButton opt_resumen 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "RESUMEN"
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
         Height          =   320
         Left            =   12360
         TabIndex        =   136
         Top             =   660
         Width           =   3495
      End
      Begin VB.OptionButton opt_kardex_linea 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "KARDEX "
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
         Height          =   320
         Left            =   7440
         TabIndex        =   132
         Top             =   2085
         Width           =   4815
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   -65880
         TabIndex        =   125
         Top             =   5280
         Width           =   6375
         Begin VitekeySoft.ChameleonBtn cmdTransacciones 
            Height          =   615
            Left            =   2640
            TabIndex        =   126
            Top             =   1680
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BTYPE           =   5
            TX              =   "GENERAR REPORTE"
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
            MICON           =   "FrmReporteProducto.frx":0EAE
            PICN            =   "FrmReporteProducto.frx":0ECA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpInicio_transacciones 
            Height          =   405
            Left            =   1440
            TabIndex        =   130
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin MSComCtl2.DTPicker Dtpfin_transacciones 
            Height          =   405
            Left            =   1440
            TabIndex        =   131
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA FINAL:"
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
            Left            =   360
            TabIndex        =   129
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INICIO :"
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
            Left            =   360
            TabIndex        =   128
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACCIONES "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   3600
            TabIndex        =   127
            Top             =   360
            Width           =   1305
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRODUCTOS VENTA DIFERIDA/PENDIENTE DE ENTREGA"
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
         Height          =   1695
         Left            =   -74280
         TabIndex        =   119
         Top             =   2520
         Width           =   5175
         Begin VitekeySoft.ChameleonBtn cmddiferida 
            Height          =   405
            Left            =   1440
            TabIndex        =   120
            ToolTipText     =   "Reporte"
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "PENDIENTE"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":34AF
            PICN            =   "FrmReporteProducto.frx":34CB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPDiferidaIni 
            Height          =   405
            Left            =   1440
            TabIndex        =   123
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
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
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin MSComCtl2.DTPicker DTPDiferidaFin 
            Height          =   405
            Left            =   3000
            TabIndex        =   139
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
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
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin VitekeySoft.ChameleonBtn cmdEntregados 
            Height          =   405
            Left            =   3000
            TabIndex        =   149
            ToolTipText     =   "Reporte"
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "ENTREGADOS"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":3558
            PICN            =   "FrmReporteProducto.frx":3574
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INICIO :"
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
            TabIndex        =   124
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame framemayor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "QUIEBRE DE STOCK"
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
         Height          =   1455
         Left            =   -74280
         TabIndex        =   117
         Top             =   960
         Width           =   5175
         Begin VB.CheckBox chk_lineaquiebre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "LINEA :"
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
            Height          =   255
            Left            =   240
            TabIndex        =   121
            Top             =   360
            Width           =   855
         End
         Begin VitekeySoft.ChameleonBtn cmdReport 
            Height          =   405
            Left            =   1200
            TabIndex        =   118
            ToolTipText     =   "Reporte"
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            BTYPE           =   3
            TX              =   "REPORTE"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":3601
            PICN            =   "FrmReporteProducto.frx":361D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcLineaQuiebre 
            Height          =   330
            Left            =   1200
            TabIndex        =   122
            Top             =   360
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
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   -65880
         TabIndex        =   113
         Top             =   2880
         Width           =   6375
         Begin VitekeySoft.ChameleonBtn cmdLiquidacionLinea 
            Height          =   615
            Left            =   2640
            TabIndex        =   114
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1085
            BTYPE           =   5
            TX              =   "GENERAR REPORTE"
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
            MICON           =   "FrmReporteProducto.frx":36AA
            PICN            =   "FrmReporteProducto.frx":36C6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblproducto 
            BackColor       =   &H00C0C0C0&
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
            Height          =   225
            Left            =   2640
            TabIndex        =   116
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK Y TOP 200 "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   2640
            TabIndex        =   115
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   -65880
         TabIndex        =   106
         Top             =   720
         Width           =   6375
         Begin MSComCtl2.DTPicker DtpLiquidacion_ini 
            Height          =   405
            Left            =   1440
            TabIndex        =   107
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin MSComCtl2.DTPicker DtpLiquidacion_fin 
            Height          =   405
            Left            =   1440
            TabIndex        =   108
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin VitekeySoft.ChameleonBtn cmdGenerarLiquidacion 
            Height          =   855
            Left            =   3240
            TabIndex        =   111
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1508
            BTYPE           =   5
            TX              =   "GENERAR REPORTE"
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
            MICON           =   "FrmReporteProducto.frx":5CAB
            PICN            =   "FrmReporteProducto.frx":5CC7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LIQUIDACION DE VENTAS"
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
            Left            =   1440
            TabIndex        =   112
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INICIO :"
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
            Left            =   360
            TabIndex        =   110
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA FINAL:"
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
            Left            =   360
            TabIndex        =   109
            Top             =   1200
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   4335
         Left            =   -74280
         TabIndex        =   87
         Top             =   1500
         Width           =   6015
         Begin VB.CheckBox chk_cliente 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "CLIENTE:"
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
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtbuscarCliente 
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
            Height          =   300
            Left            =   4920
            TabIndex        =   102
            Top             =   2040
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DtpInicio_gerencial 
            Height          =   405
            Left            =   1440
            TabIndex        =   90
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin MSComCtl2.DTPicker DtpFin_gerencial 
            Height          =   405
            Left            =   1440
            TabIndex        =   91
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin VitekeySoft.ChameleonBtn cmdGenerarReporte 
            Height          =   495
            Left            =   2040
            TabIndex        =   92
            Top             =   2760
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "GENERAR RESUMEN"
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
            BCOL            =   16576
            BCOLO           =   16576
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":82AC
            PICN            =   "FrmReporteProducto.frx":82C8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdGerencialDetallado 
            Height          =   495
            Left            =   2040
            TabIndex        =   93
            Top             =   3720
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "GENERAR DETALLADO"
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
            BCOL            =   16576
            BCOLO           =   16576
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":A8AD
            PICN            =   "FrmReporteProducto.frx":A8C9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ProgressBar progrebar_resumen 
            Height          =   195
            Left            =   2040
            TabIndex        =   94
            Top             =   2520
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar progresbardetalle 
            Height          =   195
            Left            =   2040
            TabIndex        =   95
            Top             =   3480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   344
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComCtl2.DTPicker DtpFechaReporte_gerencial 
            Height          =   405
            Left            =   1440
            TabIndex        =   97
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43227
         End
         Begin MSDataListLib.DataCombo DtcCliente 
            Height          =   330
            Left            =   1440
            TabIndex        =   101
            Top             =   2040
            Width           =   3420
            _ExtentX        =   6033
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA REPORTE:"
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
            TabIndex        =   96
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA FINAL:"
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
            Left            =   360
            TabIndex        =   89
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INICIO :"
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
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.OptionButton OptReporteMensual 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "REPORTE RESUMEN MENSUAL"
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
         Height          =   320
         Left            =   -74280
         TabIndex        =   86
         Top             =   5940
         Width           =   4095
      End
      Begin VB.OptionButton OptResumen1 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         Caption         =   "REPORTE RESUMEN GENERAL"
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
         Height          =   320
         Left            =   -74280
         TabIndex        =   85
         Top             =   1140
         Width           =   6015
      End
      Begin VB.OptionButton opt_gasto_personal 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "GASTOS PERSONAL"
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
         Height          =   320
         Left            =   -74520
         TabIndex        =   82
         Top             =   3300
         Width           =   4095
      End
      Begin VB.OptionButton opt_series_vendidas 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PRODUCTO SERIES [VENDIDAS X VENDEDOR ]"
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
         Height          =   320
         Left            =   7440
         TabIndex        =   81
         Top             =   1740
         Width           =   4815
      End
      Begin VB.OptionButton opt_series_disponibles 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PRODUCTO SERIES [DISPONIBLES ]"
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
         Height          =   320
         Left            =   7440
         TabIndex        =   80
         Top             =   1380
         Width           =   4815
      End
      Begin VB.OptionButton opt_reporte101 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VENTAS POR CLIENTE"
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
         Height          =   320
         Left            =   600
         TabIndex        =   79
         Top             =   5100
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte1013 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PRODUCTO SERIES [VENDIDAS ]"
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
         Height          =   320
         Left            =   7440
         TabIndex        =   78
         Top             =   1020
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte1012 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "REPORTE   10.12"
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
         Height          =   320
         Left            =   7440
         TabIndex        =   77
         Top             =   660
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte1011 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "REPORTE VENTAS X PRECIO [COBERTURAS]"
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
         Height          =   320
         Left            =   600
         TabIndex        =   76
         Top             =   8340
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte109 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "REPORTE   10.9"
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
         Height          =   320
         Left            =   600
         TabIndex        =   75
         Top             =   7980
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte108 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VENTAS POR VENDEDOR ARTICULO"
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
         Height          =   320
         Left            =   600
         TabIndex        =   74
         Top             =   7620
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte107 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "DESEMPEÑO X TRABAJADOR"
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
         Height          =   320
         Left            =   600
         TabIndex        =   73
         Top             =   7260
         Width           =   2775
      End
      Begin VB.OptionButton Opt_reporte106 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "UTILIDADES POR PRODUCTO"
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
         Height          =   320
         Left            =   600
         TabIndex        =   72
         Top             =   6900
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte105 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "UTILIDADES (POR PRODUCTO Y LINEA)"
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
         Height          =   320
         Left            =   600
         TabIndex        =   71
         Top             =   6540
         Width           =   3975
      End
      Begin VB.OptionButton Opt_reporte104 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "UTILIDADES (POR LINEA)"
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
         Height          =   320
         Left            =   600
         TabIndex        =   70
         Top             =   6180
         Width           =   3975
      End
      Begin VB.OptionButton Opt_reporte103 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "CLIENTES POR VENDEDOR "
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
         Height          =   320
         Left            =   600
         TabIndex        =   69
         Top             =   5820
         Width           =   4815
      End
      Begin VB.OptionButton Opt_reporte102 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VENTAS POR PRODUCTO"
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
         Height          =   320
         Left            =   600
         TabIndex        =   68
         Top             =   5460
         Width           =   4815
      End
      Begin VB.OptionButton opt_planilla_cobrador 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PLANILLA DE COBRANZA"
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
         Height          =   320
         Left            =   -74520
         TabIndex        =   67
         Top             =   2835
         Width           =   4095
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ALCANCE"
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
         Height          =   4935
         Left            =   -65640
         TabIndex        =   53
         Top             =   1680
         Width           =   6495
         Begin VB.TextBox txtCuentaContable 
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
            Left            =   2400
            TabIndex        =   84
            Top             =   3480
            Width           =   3855
         End
         Begin VB.CheckBox chk_cuenta_contable 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CUENTA CONTABLE:"
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
            Height          =   320
            Left            =   360
            TabIndex        =   83
            Top             =   3480
            Width           =   1935
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "SUCURSAL"
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
            Height          =   320
            Left            =   360
            TabIndex        =   58
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chk_cobrador 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "TRABAJADOR"
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
            Height          =   320
            Left            =   360
            TabIndex        =   56
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CLASIFICACION"
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
            Height          =   320
            Left            =   360
            TabIndex        =   55
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "MARCA"
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
            Height          =   320
            Left            =   360
            TabIndex        =   54
            Top             =   2880
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker DtpInicioRendimiento 
            Height          =   315
            Left            =   2400
            TabIndex        =   57
            Top             =   480
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
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin MSDataListLib.DataCombo DtcSucursalRendimiento 
            Height          =   330
            Left            =   2400
            TabIndex        =   59
            Top             =   1080
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
         Begin MSComCtl2.DTPicker DtpFinRendimiento 
            Height          =   315
            Left            =   4560
            TabIndex        =   60
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin MSDataListLib.DataCombo DtcvendedorRendimiento 
            Height          =   330
            Left            =   2400
            TabIndex        =   61
            Top             =   1680
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
         Begin MSDataListLib.DataCombo DtcLinearendimiento 
            Height          =   330
            Left            =   2400
            TabIndex        =   62
            Top             =   2280
            Width           =   3900
            _ExtentX        =   6879
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
         Begin MSDataListLib.DataCombo DtcMarcaRendimiento 
            Height          =   330
            Left            =   2400
            TabIndex        =   63
            Top             =   2880
            Width           =   3900
            _ExtentX        =   6879
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
         Begin VitekeySoft.ChameleonBtn cmdRendimientoDetalle 
            Height          =   375
            Left            =   2400
            TabIndex        =   104
            Top             =   3960
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "REPORTE DETALLADO"
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":DB9F
            PICN            =   "FrmReporteProducto.frx":DBBB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdInconsistencias 
            Height          =   375
            Left            =   2400
            TabIndex        =   105
            Top             =   4440
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "INCONSISTENCIAS"
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmReporteProducto.frx":FF0F
            PICN            =   "FrmReporteProducto.frx":FF2B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "RANGO FECHAS"
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
            Left            =   360
            TabIndex        =   64
            Top             =   480
            Width           =   1920
         End
      End
      Begin VB.OptionButton OptReporteventas 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "REPORTE VENTAS VENDEDOR"
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
         Height          =   320
         Left            =   -74520
         TabIndex        =   52
         Top             =   1800
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ALCANCE"
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
         Height          =   5295
         Left            =   -67560
         TabIndex        =   36
         Top             =   1560
         Width           =   8175
         Begin VB.CheckBox chk_producto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "PRODUCTO"
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
            Height          =   320
            Left            =   360
            TabIndex        =   98
            Top             =   2880
            Width           =   1815
         End
         Begin VB.CheckBox chk_sucursal_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "SUCURSAL"
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
            Height          =   320
            Left            =   360
            TabIndex        =   43
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox chk_vendedor_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "VENDEDOR"
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
            Height          =   320
            Left            =   360
            TabIndex        =   41
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chk_linea_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CLASIFICACION"
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
            Height          =   320
            Left            =   360
            TabIndex        =   40
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CheckBox chk_modelo_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "MODELO"
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
            Height          =   320
            Left            =   360
            TabIndex        =   39
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox chk_proveedor_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "PROVEEDOR"
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
            Height          =   320
            Left            =   360
            TabIndex        =   38
            Top             =   4800
            Width           =   1815
         End
         Begin VB.CheckBox chk_marca_c 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "MARCA"
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
            Height          =   320
            Left            =   360
            TabIndex        =   37
            Top             =   4080
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DtpInicio_c 
            Height          =   315
            Left            =   2400
            TabIndex        =   42
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin MSDataListLib.DataCombo DtcAlmacen_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   44
            Top             =   1080
            Width           =   4095
            _ExtentX        =   7223
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
         Begin MSComCtl2.DTPicker DtpFin_c 
            Height          =   315
            Left            =   4560
            TabIndex        =   45
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin MSDataListLib.DataCombo DtpVendedor_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   46
            Top             =   1680
            Width           =   4095
            _ExtentX        =   7223
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
         Begin MSDataListLib.DataCombo DtcLinea_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   47
            Top             =   2280
            Width           =   4140
            _ExtentX        =   7303
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
         Begin MSDataListLib.DataCombo DtcModelo_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   48
            Top             =   3480
            Width           =   4140
            _ExtentX        =   7303
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
         Begin MSDataListLib.DataCombo DtcProveedor_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   49
            Top             =   4800
            Width           =   4140
            _ExtentX        =   7303
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
         Begin MSDataListLib.DataCombo DtcMarca_c 
            Height          =   330
            Left            =   2400
            TabIndex        =   50
            Top             =   4080
            Width           =   4140
            _ExtentX        =   7303
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
         Begin MSDataListLib.DataCombo DtcProducto 
            Height          =   330
            Left            =   2400
            TabIndex        =   99
            Top             =   2880
            Width           =   4140
            _ExtentX        =   7303
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "RANGO FECHAS"
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
            Height          =   320
            Left            =   360
            TabIndex        =   51
            Top             =   480
            Width           =   1800
         End
      End
      Begin VB.OptionButton opt_venta_categoria 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "VENTAS POR CATEGORIA"
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
         Height          =   320
         Left            =   600
         TabIndex        =   33
         Top             =   4680
         Width           =   4815
      End
      Begin VB.OptionButton opt_ventas_vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "VENTAS POR VENDEDOR"
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
         Height          =   320
         Left            =   600
         TabIndex        =   32
         Top             =   4320
         Width           =   4815
      End
      Begin VB.OptionButton Opt_comparativo_diario 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMPARATIVO DIARIO DE VENTAS VALORIZADAS"
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
         Height          =   320
         Left            =   -74040
         TabIndex        =   31
         Top             =   1440
         Width           =   5055
      End
      Begin VB.OptionButton Opt_comparativo_mensual 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMPORATIVO MENSUAL DE VENTAS VALORIZADOS"
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
         Height          =   320
         Left            =   -74040
         TabIndex        =   30
         Top             =   1920
         Width           =   5055
      End
      Begin VB.OptionButton Opt_comparativo_mensual_cantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMPARATIVO MENSUAL [CANTIDADES]"
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
         Height          =   320
         Left            =   -74040
         TabIndex        =   29
         Top             =   2880
         Width           =   5055
      End
      Begin VB.OptionButton Opt_comparativo_diario_cantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMPARATIVO DIARIO [CANTIDADES]"
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
         Height          =   320
         Left            =   -74040
         TabIndex        =   28
         Top             =   2400
         Width           =   5055
      End
      Begin VB.OptionButton opt_producto_menos_vendidos 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTOS MENOS VENDIDOS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   4815
      End
      Begin VitekeySoft.ChameleonBtn cmd_reporte 
         Height          =   615
         Left            =   9960
         TabIndex        =   25
         Top             =   8100
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "GENERAR REPORTE"
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
         MICON           =   "FrmReporteProducto.frx":1227F
         PICN            =   "FrmReporteProducto.frx":1229B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ALCANCE"
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
         Left            =   7440
         TabIndex        =   9
         Top             =   2760
         Width           =   8175
         Begin VB.TextBox txtBusquedaproveedor 
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
            Height          =   300
            Left            =   7200
            TabIndex        =   162
            Top             =   4440
            Width           =   855
         End
         Begin VB.CheckBox chk_busqueda_cliente 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CLIENTE"
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
            Height          =   320
            Left            =   360
            TabIndex        =   160
            Top             =   3960
            Width           =   1815
         End
         Begin VB.TextBox txtBusquedaCliente 
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
            Height          =   300
            Left            =   7200
            TabIndex        =   156
            Top             =   3960
            Width           =   855
         End
         Begin VB.CheckBox chk_modelo_ii 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "MODELO"
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
            Height          =   320
            Left            =   360
            TabIndex        =   154
            Top             =   2580
            Width           =   1815
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2400
            TabIndex        =   140
            Top             =   4845
            Width           =   4695
            Begin VB.OptionButton opt_descripcion 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "DESCRIPCION"
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
               Height          =   195
               Left            =   2640
               TabIndex        =   142
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton opt_codigo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "CODIGO"
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
               TabIndex        =   141
               Top             =   120
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.TextBox Txtproducto 
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
            Height          =   300
            Left            =   7200
            TabIndex        =   135
            Top             =   3480
            Width           =   855
         End
         Begin VB.CheckBox chk_producto_gen 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "PRODUCTO"
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
            Height          =   320
            Left            =   360
            TabIndex        =   133
            Top             =   3480
            Width           =   1815
         End
         Begin VB.CheckBox chk_marca 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "MARCA"
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
            Height          =   320
            Left            =   360
            TabIndex        =   22
            Top             =   3000
            Width           =   1815
         End
         Begin VB.CheckBox chk_proveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "PROVEEDOR"
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
            Height          =   320
            Left            =   360
            TabIndex        =   20
            Top             =   4440
            Width           =   1815
         End
         Begin VB.CheckBox chk_modelo 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "SUB FAMILIA "
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
            Height          =   320
            Left            =   360
            TabIndex        =   18
            Top             =   2160
            Width           =   1815
         End
         Begin VB.CheckBox chk_clasificacion 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "FAMILIA"
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
            Height          =   320
            Left            =   360
            TabIndex        =   16
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chk_vendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "VENDEDOR"
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
            Height          =   320
            Left            =   360
            TabIndex        =   14
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   2400
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin VB.CheckBox chk_sucursal 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "SUCURSAL"
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
            Height          =   320
            Left            =   360
            TabIndex        =   11
            Top             =   720
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DtcAlmacen 
            Height          =   330
            Left            =   2400
            TabIndex        =   10
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSComCtl2.DTPicker DtpFin 
            Height          =   315
            Left            =   5160
            TabIndex        =   13
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   160038913
            CurrentDate     =   43096
         End
         Begin MSDataListLib.DataCombo DtcVendedor 
            Height          =   330
            Left            =   2400
            TabIndex        =   15
            Top             =   1200
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcLinea 
            Height          =   330
            Left            =   2400
            TabIndex        =   17
            Top             =   1680
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcModelo 
            Height          =   330
            Left            =   2400
            TabIndex        =   19
            Top             =   2160
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcProveedor 
            Height          =   330
            Left            =   2400
            TabIndex        =   21
            Top             =   4440
            Visible         =   0   'False
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcMarca 
            Height          =   330
            Left            =   2400
            TabIndex        =   23
            Top             =   3000
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcProductogen 
            Height          =   330
            Left            =   2400
            TabIndex        =   134
            Top             =   3480
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcModelo_ii 
            Height          =   330
            Left            =   2400
            TabIndex        =   155
            Top             =   2580
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin MSDataListLib.DataCombo DtcBusquedaCliente 
            Height          =   330
            Left            =   2400
            TabIndex        =   161
            Top             =   3960
            Visible         =   0   'False
            Width           =   4740
            _ExtentX        =   8361
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "AL"
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
            Left            =   4440
            TabIndex        =   153
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "ORDENAR POR:"
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
            Left            =   360
            TabIndex        =   143
            Top             =   4920
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "RANGO FECHAS"
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
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Width           =   1800
         End
      End
      Begin VB.OptionButton opt_cliente_menor_utilidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTES CON MENOR UTILIDAD"
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
         Height          =   320
         Left            =   600
         TabIndex        =   8
         Top             =   3960
         Width           =   4815
      End
      Begin VB.OptionButton opt_cliente_mayor_utilidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTES CON MAYOR UTILIDAD"
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
         Height          =   320
         Left            =   600
         TabIndex        =   7
         Top             =   3600
         Width           =   4815
      End
      Begin VB.OptionButton opt_clientes_menos_ventas 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTES CON MENOS VENTAS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   6
         Top             =   3240
         Width           =   4815
      End
      Begin VB.OptionButton opt_clientes_mas_ventas 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTES CON MAS VENTAS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   5
         Top             =   2880
         Width           =   4815
      End
      Begin VB.OptionButton opt_producto_menor_utilidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTOS CON MENOR UTILIDAD"
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
         Height          =   320
         Left            =   600
         TabIndex        =   4
         Top             =   2520
         Width           =   4815
      End
      Begin VB.OptionButton opt_producto_mayor_utilidad 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTOS CON MAYOR UTILIDAD"
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
         Height          =   320
         Left            =   600
         TabIndex        =   3
         Top             =   2160
         Width           =   4815
      End
      Begin VB.OptionButton opt_producto_mas_vendido 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTOS MAS VENDIDOS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   4815
      End
      Begin VB.OptionButton opt_listado_general 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "LISTA DE PRODUCTOS"
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
         Height          =   320
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   2295
      End
      Begin VitekeySoft.ChameleonBtn cmd_cerrar 
         Height          =   615
         Left            =   13560
         TabIndex        =   26
         Top             =   8100
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
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
         MICON           =   "FrmReporteProducto.frx":1486C
         PICN            =   "FrmReporteProducto.frx":14888
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdReporte_comparativo 
         Height          =   735
         Left            =   -63960
         TabIndex        =   34
         Top             =   7320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "GENERAR REPORTE"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteProducto.frx":14C78
         PICN            =   "FrmReporteProducto.frx":14C94
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar_comparativo 
         Height          =   735
         Left            =   -61680
         TabIndex        =   35
         Top             =   7320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteProducto.frx":17265
         PICN            =   "FrmReporteProducto.frx":17281
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreorterendimiento 
         Height          =   615
         Left            =   -63600
         TabIndex        =   65
         Top             =   7200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "GENERAR REPORTE"
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
         MICON           =   "FrmReporteProducto.frx":17671
         PICN            =   "FrmReporteProducto.frx":1768D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarRendimiento 
         Height          =   615
         Left            =   -61320
         TabIndex        =   66
         Top             =   7200
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
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
         MICON           =   "FrmReporteProducto.frx":19C5E
         PICN            =   "FrmReporteProducto.frx":19C7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdmensual 
         Height          =   495
         Left            =   -72600
         TabIndex        =   100
         Top             =   6420
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "GENERAR RESUMEN"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteProducto.frx":1A06A
         PICN            =   "FrmReporteProducto.frx":1A086
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDetallado 
         Height          =   615
         Left            =   7440
         TabIndex        =   137
         Top             =   8100
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "REPORTE DETALLADO"
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
         MICON           =   "FrmReporteProducto.frx":1C66B
         PICN            =   "FrmReporteProducto.frx":1C687
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar prg_avance 
         Height          =   255
         Left            =   -63600
         TabIndex        =   150
         Top             =   6840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmdUpdateCosto 
         Height          =   615
         Left            =   4680
         TabIndex        =   157
         Top             =   6225
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "UPDATE COSTO"
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
         MICON           =   "FrmReporteProducto.frx":1EC58
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar prg_reporte 
         Height          =   195
         Left            =   3480
         TabIndex        =   158
         Top             =   7320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   7860
         Left            =   -74880
         Top             =   600
         Width           =   15975
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   7500
         Left            =   -74880
         Top             =   780
         Width           =   15975
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   7380
         Left            =   -74880
         Top             =   960
         Width           =   15975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   7500
         Left            =   -74880
         Top             =   840
         Width           =   15975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   8280
         Left            =   120
         Top             =   480
         Width           =   15975
      End
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   16695
   End
End
Attribute VB_Name = "FrmReporteProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Linea As String
Dim marca As String
Public Procedencia As EnumProcede




Private Sub ChameleonBtn2_Click()


End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_busqueda_cliente_Click()


If Me.chk_busqueda_cliente.Value = 1 Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_cliente='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcBusquedaCliente)
    Me.DtcBusquedaCliente.Visible = True
Else
    Me.DtcBusquedaCliente.Visible = False
End If




End Sub

Private Sub chk_cobrador_Click()
If Me.chk_cobrador.Value = 1 Then
   Me.cmdRendimientoDetalle.Visible = True
Else
    Me.cmdRendimientoDetalle.Visible = False
End If
End Sub

Private Sub chk_linea_diferida_Click()
If chk_linea_diferida.Value = 1 Then
    strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
   ' Call LlenaDataCombo(Me.DtcLineaDiferida)
End If
End Sub

Private Sub chk_lineaquiebre_Click()
If Me.chk_lineaquiebre.Value = 1 Then
    strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLineaQuiebre)
End If
End Sub

Private Sub chk_producto_gen_Click()
If Me.chk_producto_gen.Value = 1 Then
  strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProductogen)
End If

End Sub

Private Sub chk_proveedor_Click()
If Me.chk_proveedor.Value = 1 Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_proveedor='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedor)
    Me.DtcProveedor.Visible = True
Else
    Me.DtcProveedor.Visible = False
End If
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub
Private Sub utilidad_linea()


Dim cam(0 To 1, 1 To 2)  As String
Dim cam3(0 To 2, 1 To 2)  As String

If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    strCadena = "CALL ADM_reportes_generales('4','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptUtilidadLineaProducto", param, App.Path + "\Reportes\")
    
    Exit Sub
End Sub

Private Sub utilidad_producto()

Dim cam(0 To 1, 1 To 2)  As String
Dim cam3(0 To 2, 1 To 2)  As String

If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    strCadena = "CALL ADM_reportes_generales('5','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptUtilidadProducto", param, App.Path + "\Reportes\")
    
    Exit Sub
End Sub

Private Sub persona_rendimiento(ByVal in_producto As String, ByVal in_linea As String, ByVal in_sublinea As String, ByVal in_modelo As String, ByVal in_marca As String, ByVal in_vendedor As String, ByVal in_proveedor As String)

strCadena = "DELETE FROM persona_rendimiento where dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT DISTINCT `id_vendedor` FROM movimiento_venta WHERE  id_vendedor like '%" & in_vendedor & "%' and   fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prg_reporte.Min = 0
   Me.prg_reporte.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
       
       strCadena = "CALL ADM_reportes_generales_v2('10','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          rstK.MoveFirst
          in_factura = 0
          in_boleta = 0
          in_nota = 0
          in_proforma = 0
          For j = 0 To rstK.RecordCount - 1
              Select Case rstK("id_doc")
                    Case "0001"
                        in_factura = rstK("total")
                    Case "0003"
                        in_boleta = rstK("total")
                    Case "0007"
                        in_nota = rstK("total")
                    Case "0099"
                        in_proforma = rstK("total")
                End Select
                rstK.MoveNext
         Next j
       Else
          in_factura = 0
          in_boleta = 0
          in_nota = 0
          in_proforma = 0
       End If
       strCadena = "CALL ADM_reportes_generales_v2('11','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          in_valor_venta = rstK("valor_venta")
          in_valor_total = rstK("valor_total")
       Else
          in_valor_venta = 0
          in_valor_total = 0
       End If
       
       strCadena = "CALL ADM_reportes_generales_v2('12','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          in_valor_venta_c = rstK("valor_venta")
          in_valor_total_c = rstK("valor_total")
       Else
          in_valor_venta_c = 0
          in_valor_total_c = 0
       End If
       
       strCadena = "CALL ADM_reportes_generales_v2('13','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       Call ConfiguraRstK(strCadena)
       in_clientes = rstK.RecordCount
       
       
       
       strCadena = "INSERT INTO persona_rendimiento(`id_cliente`,`ncliente`,`factura`,`boleta`,`nota`,`profoma`,clientes,`valor_venta_contado`,`valor_total_contado`,`valor_venta_credito`,`valor_total_credito`,`dni_save`,`ruc`)VALUES " & _
       "('" & rst("id_vendedor") & "','" & get_persona(rst("id_vendedor")) & "','" & in_factura & "','" & in_boleta & "','" & in_nota & "','" & in_proforma & "','" & in_clientes & "','" & in_valor_venta & "','" & in_valor_total & "','" & in_valor_venta_c & "','" & in_valor_total_c & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       Me.prg_reporte.Value = i
       rst.MoveNext
    
       DoEvents
   Next i
   
   End If
   



End Sub



Private Sub cmd_reporte_Click()




Dim in_criterio As String
Dim in_order As String

Dim cam(0 To 1, 1 To 2)  As String
Dim cam3(0 To 2, 1 To 2)  As String


If Me.opt_codigo = True Then
    in_order = " ORDER BY id_producto"
Else
    in_order = " ORDER BY nombre_prod"
End If




If Me.Opt_reporte104.Value = True Then
    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    strCadena = "CALL ADM_reportes_generales('3','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptUtilidadLinea", param, App.Path + "\Reportes\")
    Exit Sub
End If

'Ventas por Vendedor Articulo
If Me.Opt_reporte108.Value = True Then
    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    
    If Me.chk_vendedor.Value = 1 Then
       in_vendedor = Me.DtcVendedor.BoundText
    Else
       in_vendedor = ""
    End If
    
    
    If Me.chk_proveedor.Value = 1 Then
       in_proveedor = Me.DtcProveedor.BoundText
    Else
       in_proveedor = ""
    End If
    
    
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    

    
    
    strCadena = "CALL ADM_reportes_generales_v2('14','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptVentaVendedor", param, App.Path + "\Reportes\")
    
    Exit Sub
End If



If Me.Opt_reporte107.Value = True Then
    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    
    If Me.chk_vendedor.Value = 1 Then
       in_vendedor = Me.DtcVendedor.BoundText
    Else
       in_vendedor = ""
    End If
    
    
    If Me.chk_proveedor.Value = 1 Then
       in_proveedor = Me.DtcProveedor.BoundText
    Else
       in_proveedor = ""
    End If
    
    
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    Call persona_rendimiento(in_producto, in_linea, in_sublinea, in_modelo, in_marca, in_vendedor, in_proveedor)
    
    
    strCadena = "CALL ADM_reportes_generales_v2('9','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptDesempenio", param, App.Path + "\Reportes\")
    
    Exit Sub
End If





If Me.Opt_reporte105.Value = True Then
   Call utilidad_linea
   Exit Sub
End If

If Me.Opt_reporte106.Value = True Then
   Call utilidad_producto
   Exit Sub
End If






If Me.opt_listado_general.Value = True Then
    cam(0, 1) = "cambio_ini"
    cam(1, 1) = "cambio_fin"

    cam(0, 2) = str(KEY_CAMBIO_LOCAL)
    cam(1, 2) = str(KEY_CAMBIO_LOCAL)
    param = cam()

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   
   
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
    If Me.chk_producto_stock.Value = 1 Then
        If Me.chk_sucursal.Value = 1 Then
            strCadena = "SELECT '" & Me.DtcAlmacen.Text & "', id_producto,nombre_prod,linea,modelo,marca,unidad,color,precio_compra,precio_venta,precio_compra,habilitado,stock FROM view_producto WHERE stock<>0 and  id_linea <>'00009' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & in_order
        Else
            strCadena = "SELECT 'TODAS LAS SUCURSALES',id_producto,nombre_prod,linea,modelo,marca,unidad,color,precio_compra,precio_venta,precio_compra,habilitado,sum(stock) FROM  view_producto WHERE stock<>0 and id_linea <>'00009' and   ruc='" & KEY_RUC & "'  " & in_criterio & "  GROUP BY id_producto"
        End If
   Else
        If Me.chk_sucursal.Value = 1 Then
            strCadena = "SELECT '" & Me.DtcAlmacen.Text & "', id_producto,nombre_prod,linea,modelo,marca,unidad,color,precio_compra,precio_venta,precio_compra,habilitado,stock FROM view_producto WHERE  id_linea <>'00009' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & in_order
        Else
            strCadena = "SELECT 'TODAS LAS SUCURSALES',id_producto,nombre_prod,linea,modelo,marca,unidad,color,precio_compra,precio_venta,precio_compra,habilitado,sum(stock) FROM  view_producto WHERE  id_linea <>'00009' and   ruc='" & KEY_RUC & "' " & in_criterio & "  GROUP BY id_producto "
        End If
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptProducto1", param, App.Path + "\Reportes\")
   Exit Sub
End If




If Me.opt_catalogo.Value = True Then
    cam(0, 1) = "cambio_ini"
    cam(1, 1) = "cambio_fin"

    cam(0, 2) = str(KEY_CAMBIO_LOCAL)
    cam(1, 2) = str(KEY_CAMBIO_LOCAL)
    param = cam()

   
    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    
    If Me.chk_vendedor.Value = 1 Then
       in_vendedor = Me.DtcVendedor.BoundText
    Else
       in_vendedor = ""
    End If
    
    
    If Me.chk_proveedor.Value = 1 Then
       in_proveedor = Me.DtcProveedor.BoundText
    Else
       in_proveedor = ""
    End If
    
    
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = ""
    End If

    
    
    strCadena = "CALL ADM_reportes_generales_v2('15','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptProducto_precios", param, App.Path + "\Reportes\")
    Exit Sub
End If




If Me.opt_producto_mas_vendido.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_modelo_ii.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_modelo='" & Me.DtcModelo_ii.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_modelo='" & Me.DtcModelo_ii.BoundText & "'"
      End If
   
   End If
   
   
   
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,sum(cantidad),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,0,0,0,0,0,stock FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_producto,id_alm  ORDER BY 6 DESC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,sum(cantidad),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta ,0,0,0,0,0,stock      FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_producto ORDER BY 6 DESC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptProducto_rotacion", , App.Path + "\Reportes\")
   Exit Sub
End If


'PRODUCTO MENOS VENDIDO
If Me.opt_producto_menos_vendidos.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,sum(cantidad),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_producto,id_alm  ORDER BY 6 ASC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,sum(cantidad),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_producto ORDER BY 6 ASC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptProducto_rotacion", , App.Path + "\Reportes\")
   Exit Sub
End If


'PRODUCTO MAYOR UTILIDAD
If Me.opt_producto_mayor_utilidad.Value = True Then
   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,max(precio-precio_compra),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,'','','','','',sum(cantidad) FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_producto,id_alm  ORDER BY 6 DESC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,max(precio-precio_compra),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,'','','','','',sum(cantidad) FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_producto ORDER BY 6 DESC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptProducto_utilidad", , App.Path + "\Reportes\")
   Exit Sub
End If
'PRODUCTO MENOR UTILIDAD
If Me.opt_producto_menor_utilidad.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,max(precio-precio_compra),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,'','','','','',sum(cantidad) FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_producto,id_alm  ORDER BY 6 ASC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,max(precio-precio_compra),linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,'','','','','',sum(cantidad) FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_producto ORDER BY 6 ASC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptProducto_utilidad", , App.Path + "\Reportes\")
   Exit Sub
End If




'CLIENTES CON MAS VENTAS
If Me.opt_clientes_mas_ventas.Value = True Then
    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    If Me.chk_sucursal.Value = 1 Then
       in_almacen = Me.DtcAlmacen.Text
       in_alm = Me.DtcAlmacen.BoundText
    Else
       in_almacen = "TODAS LAS SUCURSALES"
       in_alm = ""
    End If
    
    If Me.chk_busqueda_cliente.Value = 1 Then
        in_cliente = Me.DtcBusquedaCliente.BoundText
    Else
        in_cliente = ""
    End If
    
    If Me.chk_proveedor.Value = 1 Then
        in_proveedor = Me.DtcProveedor.BoundText
    Else
        in_proveedor = ""
    End If
    
    
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    
    strCadena = "CALL ADM_reportes_generales_v2('20','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    
    
    Call ConfiguraRst(strCadena)
    
    Ans = ShowMultiReport(rst, "RptClientemasventa", param, App.Path + "\Reportes\")
   Exit Sub
End If




























'CLIENTES CON MENOS VENTAS
If Me.opt_clientes_menos_ventas.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(total)FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_cliente,id_alm  ORDER BY 20 asc"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(total) FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_cliente ORDER BY 20 asc"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptCliente_venta", , App.Path + "\Reportes\")
   Exit Sub
End If


'CLIENTES MAYOR UTILDIAD
If Me.opt_cliente_mayor_utilidad.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(precio-precio_compra)FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_cliente,id_alm  ORDER BY 20 DESC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(precio-precio_compra) FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_cliente ORDER BY 20 DESC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptCliente_utilidad", , App.Path + "\Reportes\")
   Exit Sub
End If

'CLIENTES menor UTILDIAD
If Me.opt_cliente_menor_utilidad.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   If Me.chk_proveedor.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(precio-precio_compra)FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY id_cliente,id_alm  ORDER BY 20 ASC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,precio,linea,marca,id_alm,id_vendedor,documento,precio,precio_compra,precio_venta,id_cliente,ncliente,'-','-',count(*),sum(precio-precio_compra) FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY id_cliente ORDER BY 20 ASC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptCliente_utilidad", , App.Path + "\Reportes\")
   Exit Sub
End If


If Me.opt_ventas_vendedor.Value = True Then
                  If Me.chk_sucursal.Value = 1 Then
                      StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
                  Else
                      StrAlmacen = ""
                  End If
               
                  operador = ""
                  
                  If Me.chk_vendedor.Value = 1 Then
                    operador = Replace(Me.DtcVendedor.BoundText, "'", "''")
                  End If
                  
                  If chk_clasificacion.Value = 1 Then
                     in_linea = DtcLinea.BoundText
                  Else
                     in_linea = ""
                  End If
                  If chk_producto_gen.Value = 1 Then
                     in_producto = DtcProductogen.BoundText
                  Else
                     in_producto = ""
                  End If
                  
                   If chk_marca.Value = 1 Then
                     in_marca = DtcMarca.BoundText
                  Else
                     in_marca = ""
                  End If
                  
                  
                  If in_linea = "" And in_producto = "" And in_marca = "" Then
                     strCadena = "SELECT `id_vendedor`,`documento`,`fecha_emision`,`ncliente`,`total`,nota,monto_nota,`nombre_completo`,`ruc` " & _
                     " FROM view_venta_vendedor WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and ruc='" & KEY_RUC & "'"
                     Call ConfiguraRst(strCadena)
                      Ans = ShowMultiReport(rst, "rpt_venta_vendedor", , App.Path + "\Reportes\")
                  Else
                        strCadena = "SELECT id_producto,nombre_prod,cantidad,precio,id_alm,fecha_emision,id_doc,serie,id_marca,marca,id_linea,linea,id_vendedor,vendedor  From " & _
                        " view_venta_detallada WHERE id_doc IN('0001','0003') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and id_linea LIKE '%" & in_linea & "%' and id_marca LIKE '%" & in_marca & "%' and id_producto LIKE  '%" & in_producto & "%' and   ruc='" & KEY_RUC & "'"
                        Call ConfiguraRst(strCadena)
                        Ans = ShowMultiReport(rst, "rpt_venta_vendedor_parametro", , App.Path + "\Reportes\")
                  End If
                  Exit Sub
End If



If Me.Opt_reporte103.Value = True Then
                  
                  in_comentario = ""
                  
                  If Me.chk_sucursal.Value = 1 Then
                      StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
                      in_comentario = Me.DtcAlmacen.Text
                  Else
                      StrAlmacen = ""
                      in_comentario = ""
                  End If
               
                  operador = ""
                  
                  If Me.chk_vendedor.Value = 1 Then
                    operador = Replace(Me.DtcVendedor.BoundText, "'", "''")
                    in_comentario = in_comentario & " " & Me.DtcVendedor.Text
                  End If
                  
                  If chk_clasificacion.Value = 1 Then
                     in_linea = DtcLinea.BoundText
                     in_comentario = in_comentario & " : " & Me.DtcLinea.Text
                  Else
                     in_linea = ""
                  End If
                  If chk_producto_gen.Value = 1 Then
                     in_producto = DtcProductogen.BoundText
                     PP = Me.DtcProductogen.Text
                     in_comentario = in_comentario & " " & Me.DtcProductogen.Text
                  Else
                     in_producto = ""
                     PP = ""
                  End If
                  
                   If chk_marca.Value = 1 Then
                     in_marca = DtcMarca.BoundText
                     in_comentario = in_comentario & " " & Me.DtcMarca.Text
                  Else
                     in_marca = ""
                  End If
                 If Me.chk_proveedor.Value = 1 Then
                     in_proveedor = Me.DtcProveedor.BoundText
                     in_comentario = in_comentario & " " & Me.DtcMarca.Text
                  Else
                     in_proveedor = ""
                  End If
                  
                  
                  
                    cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
                    cam3(2, 2) = Me.DtcProveedor.Text
                    param = cam3()
                  
                  
                  
                  
                       strCadena = "SELECT COUNT(DISTINCT id_venta),vendedor,sum(precio*cantidad)  From " & _
                       " view_venta_detallada WHERE  id_doc IN('0001','0003') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_proveedor LIKE '%" & in_proveedor & "%' and  id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and id_linea LIKE '%" & in_linea & "%' and id_marca LIKE '%" & in_marca & "%' and id_producto LIKE  '%" & in_producto & "%' and   ruc='" & KEY_RUC & "' GROUP BY  id_vendedor,ruc"
                        Call ConfiguraRst(strCadena)
                        
                   
                      Ans = ShowMultiReport(rst, "rpt_venta_vendedor_cliente", param, App.Path + "\Reportes\")
                 
                  Exit Sub
End If













If Me.opt_venta_categoria.Value = True Then

    If Me.chk_producto.Value = 1 Then
       in_producto = Me.DtcProducto.BoundText
    Else
       in_producto = ""
    End If
    If Me.chk_clasificacion.Value = 1 Then
       in_linea = Me.DtcLinea.BoundText
    Else
       in_linea = ""
    End If
    If Me.chk_modelo.Value = 1 Then
       in_sublinea = Me.DtcModelo.BoundText
    Else
       in_sublinea = ""
    End If
    If Me.chk_modelo_ii.Value = 1 Then
       in_modelo = Me.DtcModelo_ii.BoundText
    Else
       in_modelo = ""
    End If
    
    If Me.chk_marca.Value = 1 Then
       in_marca = Me.DtcMarca.BoundText
    Else
       in_marca = ""
    End If
    
    If Me.chk_vendedor.Value = 1 Then
       in_vendedor = Me.DtcVendedor.BoundText
    Else
       in_vendedor = ""
    End If
    
    
    If Me.chk_proveedor.Value = 1 Then
       in_proveedor = Me.DtcProveedor.BoundText
    Else
       in_proveedor = ""
    End If
    
    
    If Me.chk_sucursal.Value = 1 Then
       in_alm = Me.DtcAlmacen.Text
    Else
       in_alm = "TODAS LAS SUCURSALES"
    End If
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
    

    
    
    strCadena = "CALL ADM_reportes_generales_v2('16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptVentaCategoria", param, App.Path + "\Reportes\")
    
    
    
    
    
    If MsgBox("Desea Visualizar el Detalle de Comprobante", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
        strCadena = "CALL ADM_reportes_generales_v2('17','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptVentaCategoria", param, App.Path + "\Reportes\")
    End If
    
    Exit Sub
   
End If









































If Me.opt_reporte101.Value = True Then
                    cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
                   
                   
                    
   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_linea = Me.DtcLinea.BoundText
   Else
      in_linea = ""
   End If
   
   
   If Me.chk_marca.Value = 1 Then
      in_marca = Me.DtcMarca.BoundText
   Else
      in_marca = ""
   End If
   
   If Me.chk_modelo.Value = 1 Then
      in_subfamilia = Me.DtcModelo.BoundText
    Else
      in_subfamilia = ""
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
      in_producto = Me.DtcProductogen.BoundText
   Else
      in_producto = ""
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      in_proveedor = DtcProveedor.BoundText
   Else
      in_proveedor = ""
   End If
   
   If Me.chk_busqueda_cliente.Value = 1 Then
      in_cliente = Me.DtcBusquedaCliente.BoundText
   Else
      in_cliente = ""
   End If
   
   
   
   
   If Me.chk_sucursal.Value = 1 Then
        in_almacen = Me.DtcAlmacen.BoundText
         cam3(2, 2) = Me.DtcAlmacen.Text
   Else
        in_almacen = ""
        If in_cliente <> "" Then
           cam3(2, 2) = "TODOS LAS SUCURSALES" & Space(2) & in_cliente & Space(2) & get_persona(in_cliente)
        End If
         
        If in_producto <> "" Then
           cam3(2, 2) = "TODOS LAS SUCURSALES" & Space(2) & in_producto & Space(2) & get_producto(in_producto)
        End If
        
   End If
   param = cam3()
   
   strCadena = "call ADM_reportes_generales('1','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_subfamilia & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "Rpt101", param, App.Path + "\Reportes\")
   
   Exit Sub
End If

'SERIALES

If Me.opt_series_disponibles.Value = True Then

   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
   
   
   
    
    cam3(0, 1) = "inicial"
    cam3(1, 1) = "final"
    cam3(2, 1) = "almacen"
    
    
   If Me.chk_sucursal.Value = 1 Then
        cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
        cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
        cam3(2, 2) = Me.DtcAlmacen.Text & Space(5) & "OPERADOR:" & KEY_VENDEDOR
        strCadena = "SELECT id_detalle,id_producto,nombre_prod,nro_chasis,nro_motor,linea,modelo,color,id_linea,id_sublinea FROM view_producto_seriales WHERE vendido='no' and transferencia='no' and nro_chasis<>'' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & in_order
   Else
        cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
        cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
        cam3(2, 2) = "TODAS LA SUCURSALES" & Chr(13) & "OPERADOR   :" & KEY_VENDEDOR
        strCadena = "SELECT id_detalle,id_producto,nombre_prod,nro_chasis,nro_motor,linea,modelo,color,id_linea,id_sublinea FROM view_producto_seriales WHERE vendido='no' and transferencia='no' and nro_chasis<>'' and   ruc='" & KEY_RUC & "' " & in_criterio & in_order
   End If
    param2 = cam3()
   Call ConfiguraRst(strCadena)
   
   

   Ans = ShowMultiReport(rst, "Rptreporte_seriales", param2, App.Path + "\Reportes\")
   
   End If
   
   
   
   
If Me.opt_series_vendidas.Value = True Then
    in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
    
    
    
    Dim cam2(0 To 2, 1 To 2)  As String
    cam2(0, 1) = "inicial"
    cam2(1, 1) = "final"
    cam2(2, 1) = "almacen"
    
    cam2(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam2(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    
        
        
        
    
    
    If Me.chk_sucursal.Value = 1 Then
        cam2(2, 2) = Me.DtcAlmacen.Text
        strCadena = "SELECT id_venta,documento,fecha_emision,id_cliente,ncliente,total,id_vendedor,nombre_completo,id_linea,descripcion,nro_chasis,serie " & _
        " FROM view_venta_produccion WHERE nro_chasis<>'-' and   id_alm =  '" & Me.DtcAlmacen.BoundText & "' and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "' " & in_criterio
    Else
        strCadena = "SELECT id_venta,documento,fecha_emision,id_cliente,ncliente,total,id_vendedor,nombre_completo,id_linea,descripcion,nro_chasis,serie " & _
        " FROM view_venta_produccion WHERE nro_chasis<>'-'  and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "' " & in_criterio
        cam2(2, 2) = "TODAS LAS SUCURSALES"
    End If
   
   param = cam2
   
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "rpt_ventas_produccion", param, App.Path + "\Reportes\")
   Exit Sub
   End If






If Opt_reporte1013.Value = True Then
    in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca.BoundText & "'"
      End If
   
   End If
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
   
   End If
    
    
    
    
    cam2(0, 1) = "inicial"
    cam2(1, 1) = "final"
    cam2(2, 1) = "almacen"
    
    cam2(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam2(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    
        
        
        
    
    
    If Me.chk_sucursal.Value = 1 Then
        cam2(2, 2) = Me.DtcAlmacen.Text
        strCadena = "SELECT id_venta,documento,fecha_emision,id_cliente,ncliente,total,id_vendedor,nombre_completo,id_linea,descripcion,nro_chasis,serie " & _
        " FROM view_venta_produccionv2 WHERE    id_alm =  '" & Me.DtcAlmacen.BoundText & "' and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "' " & in_criterio
    Else
        strCadena = "SELECT id_venta,documento,fecha_emision,id_cliente,ncliente,total,id_vendedor,nombre_completo,id_linea,descripcion,nro_chasis,serie " & _
        " FROM view_venta_produccionv2 WHERE   fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "' " & in_criterio
        cam2(2, 2) = "TODAS LAS SUCURSALES"
    End If
   
   
   param = cam2
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "rpt_ventas_produccion_produccion", param, App.Path + "\Reportes\")
   Exit Sub
   End If








'Obsequios
If Me.opt_obsequios.Value = True Then
   
    
    
   
    cam2(0, 1) = "inicial"
    cam2(1, 1) = "final"
    cam2(2, 1) = "almacen"
    
    cam2(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam2(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    
    If Me.chk_proveedor.Value = 1 Then
       in_proveedor = Me.DtcProveedor.BoundText
    Else
       in_proveedor = ""
    End If
        
        
    
    
   
        strCadena = "SELECT id_venta,fecha_emision,documento,id_producto,cantidad,detalle,precio,descuento,descuento_porcentaje,total,obsequio,ruc " & _
        " FROM view_obsequios_bonificaciones WHERE id_proveedor LIKE '%" & in_proveedor & "%' and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "' "
        cam2(2, 2) = "TODAS LAS SUCURSALES"
  
   
   param = cam2
   
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptObsequiosBonificaciones", param, App.Path + "\Reportes\")
   Exit Sub
   End If

' End Obsequios






   
   
   
   
   'REPORTE 10.2
   If Me.Opt_reporte102.Value = True Then

      
                   
                   
                    
   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_linea = Me.DtcLinea.BoundText
   Else
      in_linea = ""
   End If
   
   
   If Me.chk_marca.Value = 1 Then
      in_marca = Me.DtcMarca.BoundText
   Else
      in_marca = ""
   End If
   
   If Me.chk_modelo.Value = 1 Then
      in_subfamilia = Me.DtcModelo.BoundText
    Else
      in_subfamilia = ""
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
      in_producto = Me.DtcProductogen.BoundText
   Else
      in_producto = ""
   End If
   
   
   If Me.chk_proveedor.Value = 1 Then
      in_cliente = DtcProveedor.BoundText
   Else
      in_cliente = ""
   End If
   
   
   If Me.chk_sucursal.Value = 1 Then
        in_almacen = Me.DtcAlmacen.BoundText
         cam3(2, 2) = Me.DtcAlmacen.Text
   Else
        in_almacen = ""
         
         
          If in_cliente <> "" Then
           cam3(2, 2) = "TODOS LAS SUCURSALES" & Space(2) & in_cliente & Space(2) & get_persona(in_cliente)
        End If
         
        If in_producto <> "" Then
           cam3(2, 2) = "TODOS LAS SUCURSALES" & Space(2) & in_producto & Space(2) & get_producto(in_producto)
        End If
         
         
         
   End If
   cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_alm
    param = cam3()
   param = cam3()
   
   strCadena = "CALL ADM_reportes_generales_v2('19','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_cliente & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
   
   'strCadena = "call ADM_reportes_generales('8','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_subfamilia & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptVentaProducto", param, App.Path + "\Reportes\")
   
   Exit Sub
End If
   
   
   'REPORTE KARDEX
   If Me.opt_kardex_linea.Value = True Then
   in_criterio = ""
   If Me.chk_clasificacion.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   
   If Me.chk_producto_gen.Value = 1 Then
        If in_criterio <> "" Then
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_producto='" & Me.DtcProductogen.BoundText & "'"
      End If
   End If
   
   If Me.chk_modelo.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo.BoundText & "'"
      End If
  End If
   

   
Dim arr(0 To 1, 1 To 2)   As String

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "   -   " & Format(Me.DtpFin.Value, "dd-mm-YYYY")
arr(1, 2) = Me.DtcLinea.Text

param = arr()

strCadena = "DELETE FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  k.fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' and p.id_linea='" & DtcLinea.BoundText & "' ORDER BY k.id_producto"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   
   strCadena = "call put_crear_kardex_temporal('" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "');"
   CnBd.Execute (strCadena)
   
   
   For i = 0 To rst.RecordCount - 1
      strCadena = "CALL procedure_kardex_general_linea('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      rst.MoveNext
     
      Me.cmd_reporte.Caption = (i + 1) & Space(2) & "-" & Space(1) & rst.RecordCount
       DoEvents
   Next i

End If
strCadena = "SELECT id_producto,producto,cantidad_inicial,saldo_inicial,cantidad_ingreso,saldo_ingreso,cantidad_salida,saldo_salida,cantidad_final,saldo_final FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptKardexValorizado_", param, App.Path + "\Reportes\")
Exit Sub
End If







If Me.Opt_reporte1011.Value = True Then
                  
                  in_comentario = ""
                  
                  If Me.chk_sucursal.Value = 1 Then
                      StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
                      in_comentario = Me.DtcAlmacen.Text
                  Else
                      StrAlmacen = ""
                      in_comentario = ""
                  End If
               
                  operador = ""
                  
                  If Me.chk_vendedor.Value = 1 Then
                    operador = Replace(Me.DtcVendedor.BoundText, "'", "''")
                    in_comentario = in_comentario & " " & Me.DtcVendedor.Text
                  End If
                  
                  If chk_clasificacion.Value = 1 Then
                     in_linea = DtcLinea.BoundText
                     in_comentario = in_comentario & " : " & Me.DtcLinea.Text
                  Else
                     in_linea = ""
                  End If
                  If chk_producto_gen.Value = 1 Then
                     in_producto = DtcProductogen.BoundText
                     PP = Me.DtcProductogen.Text
                     in_comentario = in_comentario & " " & Me.DtcProductogen.Text
                  Else
                     in_producto = ""
                     PP = ""
                  End If
                  
                   If chk_marca.Value = 1 Then
                     in_marca = DtcMarca.BoundText
                     in_comentario = in_comentario & " " & Me.DtcMarca.Text
                  Else
                     in_marca = ""
                  End If
                  
                  
                  If Me.chk_proveedor.Value = 1 Then
                    in_proveedor = Me.DtcProveedor.BoundText
                  Else
                    in_proveedor = ""
                  End If
                  
                    cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
                    cam3(2, 2) = in_comentario
                    param = cam3()
                  
                  
                  
                  
                       strCadena = "SELECT `fecha_emision`,`id_producto`,`nombre_prod`,`unidad`,`linea`,`cantidad`,`precio`,`cobertura`,`mayor`,`mercado`,`id_alm`,`id_vendedor`,`nombre_completo`,`id_tipo_cobertura`,`tipo_cobertura`,`ruc`  From " & _
                       " view_ventasxpreciocobertura WHERE id_proveedor LIKE '%" & in_proveedor & "%' and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and id_linea LIKE '%" & in_linea & "%' and id_producto LIKE  '%" & in_producto & "%' and   ruc='" & KEY_RUC & "'"
                        Call ConfiguraRst(strCadena)
                        
                   
                      Ans = ShowMultiReport(rst, "RptVentasPrecioCobertura", param, App.Path + "\Reportes\")
                 
                  Exit Sub
End If




Exit Sub







End Sub




Private Sub cmdCerrarRendimiento_Click()
Unload Me
End Sub
Private Sub put_generar_resumen()


strCadena = "DELETE FROM reporte_cuentas_cobrar_resumen WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM cuentas_cobrar_parametros WHERE fin<>0 ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.progrebar_resumen.Min = 0
   Me.progrebar_resumen.Max = rst.RecordCount - 1
   For i = 0 To rst.RecordCount - 1
   
        
        
        'FACTURADO
        'FACTURADO
     '   strCadena = "SELECT `v`.`id_venta`,  sum( if((`v`.`id_doc` = '0007'),(`v`.`total` * -(1)),`v`.`total`) AS `total`, sum(if((`v`.`id_doc` = '0007'),(`function_pago_factura`(`v`.`id_venta`,'" & Format(Me.DtpFechaReporte_gerencial.Value, "YYYY-mm-dd") & "',`v`.`id_moneda`,`v`.`ruc`) * -(1)),`function_pago_factura`(`v`.`id_venta`,'" & Format(Me.DtpFechaReporte_gerencial.Value, "YYYY-mm-dd") & "',`v`.`id_moneda`,`v`.`ruc`)))) AS `pago`,(to_days('" & Format(Me.DtpFechaReporte_gerencial.Value, "YYYY-mm-dd") & "') - to_days(`v`.`fecha_vencimiento`)) AS `dias_vencidos`,`v`.`documento` AS `documento`, " & _
        "`v`.`fecha_vencimiento` AS `fecha_vencimiento`,`v`.`fecha_emision` AS `fecha_emision`,`v`.`id_cliente` AS `id_cliente`, " & _
        " `v`.`ncliente` AS `ncliente`,`v`.`id_doc` AS `id_doc`,`v`.`cobranza_dudosa` AS `cobranza_dudosa`, " & _
        "`v`.`fecha_cobranza_dudosa` AS `fecha_cobranza_dudosa`,`v`.`flag_saldo` AS `flag_saldo`, `v`.`ruc` AS `ruc` From  `movimiento_venta` `v` Where  ((`v`.`id_forma_pago` = '02') and (`v`.`id_doc` in ('0001','0003','0412','0000','0007','0008')))  Order By     (to_days('" & Format(Me.DtpFechaReporte_gerencial.Value, "YYYY-mm-dd") & "') - to_days(`v`.`fecha_vencimiento`)) "
        If rst("fin") < 0 Then
            If Me.chk_cliente.Value = 1 Then
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and cobranza_dudosa='no' and  id_forma_pago='02' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos<0 and ruc='" & KEY_RUC & "'"
            Else
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE cobranza_dudosa='no' and  id_forma_pago='02' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos<0  and ruc='" & KEY_RUC & "'"
            End If
        Else
            If Me.chk_cliente.Value = 1 Then
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and cobranza_dudosa='no' and  id_forma_pago='02' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos>='" & rst("inicio") & "' and dias_vencidos<='" & rst("fin") & "' and ruc='" & KEY_RUC & "'"
            Else
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE cobranza_dudosa='no' and  id_forma_pago='02' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos>='" & rst("inicio") & "' and dias_vencidos<='" & rst("fin") & "' and ruc='" & KEY_RUC & "'"
            End If
        
        End If
        
        
        
        
        Call ConfiguraRstZ(strCadena)
        If IsNull(rstZ(0)) = True Then
           in_facturado = 0
        Else
           in_facturado = rstZ(0)
        End If
        
        If IsNull(rstZ(1)) = True Then
           in_pagado = 0
        Else
           in_pagado = rstZ(1)
        End If
        If IsNull(rstZ(2)) = True Then
           in_clientes = 0
        Else
           in_clientes = rstZ(2)
        End If
        
        in_saldo = in_facturado - in_pagado
        
        
        
        strCadena = "INSERT INTO reporte_cuentas_cobrar_resumen(descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc,dni_save,ruc)VALUES " & _
        "('" & rst("descripcion") & "','" & in_facturado & "','" & in_pagado & "','" & in_saldo & "','" & in_saldo / KEY_CAMBIO_COMPRA & "','" & in_clientes & "','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Me.progrebar_resumen.Value = i
        DoEvents
        rst.MoveNext
   Next i
   
        
        
        
        strCadena = "SELECT * FROM cuentas_cobrar_parametros where fin=0 ORDER BY id ASC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            If Me.chk_cliente.Value = 1 Then
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and cobranza_dudosa='si' and ruc='" & KEY_RUC & "' "
            Else
                strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and cobranza_dudosa='si' and ruc='" & KEY_RUC & "' "
            End If
            Call ConfiguraRstZ(strCadena)
            If IsNull(rstZ(0)) = True Then
               in_facturado = 0
            Else
               in_facturado = rstZ(0)
            End If
        
            If IsNull(rstZ(1)) = True Then
               in_pagado = 0
            Else
               in_pagado = rstZ(1)
            End If
            If IsNull(rstZ(2)) = True Then
               in_clientes = 0
            Else
               in_clientes = rstZ(2)
            End If
        
            in_saldo = in_facturado - in_pagado
            
            strCadena = "INSERT INTO reporte_cuentas_cobrar_resumen(descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc,dni_save,ruc)VALUES " & _
            "('" & rst("descripcion") & "','" & in_facturado & "','" & in_pagado & "','" & in_saldo & "','" & in_saldo / KEY_CAMBIO_COMPRA & "','" & in_clientes & "','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
    End If
    
    
    strCadena = "SELECT * FROM cuentas_cobrar_parametros where fin=0 ORDER BY id DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
         If Me.chk_cliente.Value = 1 Then
            strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and  fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and cobranza_dudosa='si' and ruc='" & KEY_RUC & "' "
         Else
            strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and cobranza_dudosa='si' and ruc='" & KEY_RUC & "' "
         End If
        Call ConfiguraRstZ(strCadena)
        If IsNull(rstZ(0)) = True Then
           in_facturado = 0
        Else
           in_facturado = rstZ(0)
        End If
        
        If IsNull(rstZ(1)) = True Then
           in_pagado = 0
        Else
           in_pagado = rstZ(1)
        End If
        If IsNull(rstZ(2)) = True Then
           in_clientes = 0
        Else
           in_clientes = rstZ(2)
        End If
        
        in_saldo = in_facturado - in_pagado
        
        
        
        strCadena = "INSERT INTO reporte_cuentas_cobrar_resumen(descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc,dni_save,ruc)VALUES " & _
        "('" & rst("descripcion") & "','" & in_facturado * -1 & "','" & in_pagado * -1 & "','" & in_saldo * -1 & "','" & (in_saldo / KEY_CAMBIO_COMPRA) * -1 & "','" & in_clientes & "','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        
    End If
   
End If






End Sub

Private Sub put_generar_resumen_mensual()


strCadena = "DELETE FROM reporte_cuentas_cobrar_resumen WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM mes WHERE id_mes<='" & Month(KEY_FECHA) & "' ORDER BY id_mes ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
            
       'acumular
        'strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle WHERE fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos>='" & rst("inicio") & "' and dias_vencidos<='" & rst("fin") & "' and ruc='" & KEY_RUC & "'"
        strCadena = "SELECT sum(total),sum(pago),COUNT(*) FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_forma_pago='02' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and month(fecha_emision)='" & rst("id_mes") & "' and ruc='" & KEY_RUC & "' "
        Call ConfiguraRstZ(strCadena)
        If IsNull(rstZ(0)) = True Then
           in_facturado = 0
        Else
           in_facturado = rstZ(0)
        End If
        
        If IsNull(rstZ(1)) = True Then
           in_pagado = 0
        Else
           in_pagado = rstZ(1)
        End If
        If IsNull(rstZ(2)) = True Then
           in_clientes = 0
        Else
           in_clientes = rstZ(2)
        End If
        
        in_saldo = in_facturado - in_pagado
        
        
        
        strCadena = "INSERT INTO reporte_cuentas_cobrar_resumen(descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc,dni_save,ruc)VALUES " & _
        "('" & rst("descripcion") & "','" & in_facturado & "','" & in_pagado & "','" & in_saldo & "','" & in_saldo / KEY_CAMBIO_COMPRA & "','" & in_clientes & "','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If






End Sub


Private Sub cmdDetallado_Click()

'modificar temp






Dim cam3(0 To 2, 1 To 2)  As String
 in_comentario = ""
                  
                  If Me.chk_sucursal.Value = 1 Then
                      StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
                      in_comentario = Me.DtcAlmacen.Text
                  Else
                      StrAlmacen = ""
                      in_comentario = ""
                  End If
               
                  operador = ""
                  
                  If Me.chk_vendedor.Value = 1 Then
                    operador = Replace(Me.DtcVendedor.BoundText, "'", "''")
                    in_comentario = in_comentario & " " & Me.DtcVendedor.Text
                  End If
                  
                  If chk_clasificacion.Value = 1 Then
                     in_linea = DtcLinea.BoundText
                     in_comentario = in_comentario & " : " & Me.DtcLinea.Text
                  Else
                     in_linea = ""
                  End If
                  If chk_producto_gen.Value = 1 Then
                     in_producto = DtcProductogen.BoundText
                     PP = Me.DtcProductogen.Text
                     in_comentario = in_comentario & " " & Me.DtcProductogen.Text
                  Else
                     in_producto = ""
                     PP = ""
                  End If
                  
                   If chk_marca.Value = 1 Then
                     in_marca = DtcMarca.BoundText
                     in_comentario = in_comentario & " " & Me.DtcMarca.Text
                  Else
                     in_marca = ""
                  End If
                  
                    cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
                    cam3(2, 2) = in_comentario
                    param = cam3()
                  
                  
                  
                  
                       strCadena = "SELECT fecha_emision,documento,ncliente,nombre_prod,cantidad,precio,cantidad*precio,vendedor  From " & _
                       " view_venta_detallada WHERE id_doc IN('0001','0003') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & StrAlmacen & "%' and id_vendedor LIKE '%" & operador & "%' and id_linea LIKE '%" & in_linea & "%' and id_marca LIKE '%" & in_marca & "%' and id_producto LIKE  '%" & in_producto & "%' and   ruc='" & KEY_RUC & "' "
                       Call ConfiguraRst(strCadena)
                       Ans = ShowMultiReport(rst, "rpt_venta_vendedor_cliente_detallada", param, App.Path + "\Reportes\")
                 
                  Exit Sub
End Sub

Private Sub cmddiferida_Click()


strCadena = "DELETE FROM movimiento_venta_diferida WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)




strCadena = "call cursor_diferida('" & Format(Me.DTPDiferidaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DTPDiferidaFin.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT `id_venta`,`fecha_emision`,`id_producto`,`detalle`,`linea`,`documento`,`id_cliente`,`ncliente`,`precio_costo`,`precio`,`cantidad_entregar`,nombre_completo FROM view_diferida13 WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_diferidas", , App.Path + "\Reportes\")


End Sub

Private Sub cmdEntregados_Click()
strCadena = "DELETE FROM movimiento_venta_diferida WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "call cursor_diferida_entregados('" & Format(Me.DTPDiferidaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DTPDiferidaFin.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


strCadena = "SELECT `id_venta`,`fecha_emision`,fecha_entrega,`id_producto`,detalle,`linea`,`documento`,guia,`id_cliente`,`ncliente`,`precio_costo`,`precio`,`cantidad_entregar`,nombre_completo FROM view_diferida_entregado_report WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_diferidas_entregado", , App.Path + "\Reportes\")
End Sub

Private Sub cmdEstadoCuentaGeneral_Click()

If chk_personal.Value = 1 Then

    strCadena = "SELECT fecha_emision,fecha_vencimiento,id_cliente,nombre_completo,direccion,observacion,celular,documento,total,monto_pagado,dias_vencidos,ubigeo,id_personal FROM view_estado_cuenta WHERE id_personal='si' and  (total-monto_pagado)>0.1 and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT fecha_emision,fecha_vencimiento,id_cliente,nombre_completo,direccion,observacion,celular,documento,total,monto_pagado,dias_vencidos,ubigeo,id_personal FROM view_estado_cuenta WHERE (total-monto_pagado)>0.1 and ruc='" & KEY_RUC & "'"
End If
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rptEstadoCuentaCliente", , App.Path + "\Reportes\")
End Sub

Private Sub cmdGenerarLiquidacion_Click()

strCadena = "SELECT id_detalle_venta,fecha_emision,hora,id_vendedor,nombre_completo,id_doc,numero,id_cliente,ncliente,id_producto,nombre_prod,descripcion,linea,modelo,unidad,marca,cantidad,precio_sistema,precio_venta,precio_costo FROM view_liquidacion_venta WHERE fecha_emision>='" & Format(Me.DtpLiquidacion_ini.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpLiquidacion_fin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_liquidacion_venta", , App.Path + "\Reportes\")

'fecha_emision,hora,id_vendedor,nombre_completo,documento,id_cliente,ncliente,id_producto,nombre_prod, " & _
" descripcion,linea,modelo,unidad,cantidad,precio_sistema,precio_venta,precio_costo,0," & _
"0,ruc


End Sub

Private Sub cmdgenerarreporte_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"



arr(0, 2) = Format(Me.DtpInicio_gerencial.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin_gerencial.Value, "dd-mm-YYYY")


param = arr()


If OptResumen1.Value = True Then
    Call put_generar_resumen
    strCadena = "SELECT id,descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc FROM reporte_cuentas_cobrar_resumen WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptResumen_credito", param, App.Path + "\Reportes\")
End If



If Me.OptReporteMensual.Value = True Then
    Call put_generar_resumen_mensual
    strCadena = "SELECT id,descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc FROM reporte_cuentas_cobrar_resumen WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptResumen_credito_mensual", , App.Path + "\Reportes\")
End If


End Sub

Private Sub cmdGerencialDetallado_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "fecha_ini"
arr(1, 1) = "fecha_fin"



arr(0, 2) = Format(Me.DtpInicio_gerencial.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin_gerencial.Value, "dd-mm-YYYY")


param = arr()

'strCadena = "UPDATE movimiento_venta SET flag_saldo=0 WHERE flag_saldo<>'0' and ruc='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)
strCadena = "DELETE FROM cuentas_cobrar_detalle WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' "
CnBd.Execute (strCadena)




strCadena = "SELECT * FROM cuentas_cobrar_parametros WHERE fin<>'0' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.progresbardetalle.Min = 0
   Me.progresbardetalle.Max = rst.RecordCount - 1
   For i = 0 To rst.RecordCount - 1
       
       If rst("fin") < 0 Then
            If Me.chk_cliente.Value = 1 Then
                 strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and  id_forma_pago='02' and cobranza_dudosa='no' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos<0  and ruc='" & KEY_RUC & "'"
            Else
                 strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_forma_pago='02' and cobranza_dudosa='no' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos<0 and ruc='" & KEY_RUC & "'"
            End If
       Else
            If Me.chk_cliente.Value = 1 Then
                 strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and  id_forma_pago='02' and cobranza_dudosa='no' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos>='" & rst("inicio") & "' and dias_vencidos<='" & rst("fin") & "' and ruc='" & KEY_RUC & "'"
            Else
                 strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_forma_pago='02' and cobranza_dudosa='no' and  fecha_emision>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and dias_vencidos>='" & rst("inicio") & "' and dias_vencidos<='" & rst("fin") & "' and ruc='" & KEY_RUC & "'"
            End If
       End If
       
       
       
       
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount > 0 Then
          rstL.MoveFirst
          For j = 0 To rstL.RecordCount - 1
              
               
               strCadena = "INSERT INTO cuentas_cobrar_detalle(id_parametro,descripcion,emision,vencimiento,dni,cliente,documento,total,pago,dudosa,id_doc,dni_save,ruc)VALUES " & _
               "('" & rst("id") & "','" & rst("descripcion") & "','" & Format(rstL("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstL("fecha_vencimiento"), "YYYY-mm-dd") & "','" & rstL("id_cliente") & "','" & rstL("ncliente") & "','" & rstL("documento") & "','" & rstL("total") & "','" & rstL("pago") & "','no','" & rstL("id_doc") & "','" & KEY_USUARIO & "','" & rstL("ruc") & "')"
               CnBd.Execute (strCadena)
               
               
               'strCadena = "UPDATE movimiento_venta SET flag_saldo='" & rst("id") & "' WHERE id_venta='" & rstL("id_venta") & "'"
               'CnBd.Execute (strCadena)
               rstL.MoveNext
          Next j
       End If
       DoEvents
       Me.progresbardetalle.Value = i
       rst.MoveNext
   Next i
   
       
       strCadena = "SELECT * FROM cuentas_cobrar_parametros WHERE fin=0 ORDER BY id ASC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
       If Me.chk_cliente.Value = 1 Then
            strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and  cobranza_dudosa='si' and  fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
       Else
            strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE cobranza_dudosa='si' and  fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
       End If
       Call ConfiguraRstL(strCadena)
           If rstL.RecordCount > 0 Then
                rstL.MoveFirst
                    For j = 0 To rstL.RecordCount - 1
                        strCadena = "INSERT INTO cuentas_cobrar_detalle(id_parametro,descripcion,emision,vencimiento,dni,cliente,documento,total,pago,dudosa,id_doc,dni_save,ruc)VALUES " & _
                        "('" & rst("id") & "','" & rst("descripcion") & "','" & Format(rstL("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstL("fecha_vencimiento"), "YYYY-mm-dd") & "','" & rstL("id_cliente") & "','" & rstL("ncliente") & "','" & rstL("documento") & "','" & rstL("total") & "','" & rstL("pago") & "','si','" & rstL("id_doc") & "','" & KEY_USUARIO & "','" & rstL("ruc") & "')"
                        CnBd.Execute (strCadena)
                        rstL.MoveNext
                    Next j
            End If
            
        End If
        
        
        
        strCadena = "SELECT * FROM cuentas_cobrar_parametros WHERE fin=0 ORDER BY id DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        If Me.chk_cliente.Value = 1 Then
            strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE id_cliente='" & Me.DtcCliente.BoundText & "' and  cobranza_dudosa='si' and  fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT * FROM view_cuentas_cobrar_resumen_detalle_v2 WHERE cobranza_dudosa='si' and  fecha_cobranza_dudosa>='" & Format(Me.DtpInicio_gerencial.Value, "YYYY-mm-dd") & "' and fecha_cobranza_dudosa<='" & Format(Me.DtpFin_gerencial.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
        End If
       Call ConfiguraRstL(strCadena)
           If rstL.RecordCount > 0 Then
                rstL.MoveFirst
                    For j = 0 To rstL.RecordCount - 1
                        strCadena = "INSERT INTO cuentas_cobrar_detalle(id_parametro,descripcion,emision,vencimiento,dni,cliente,documento,total,pago,dudosa,id_doc,dni_save,ruc)VALUES " & _
                        "('" & rst("id") & "','" & rst("descripcion") & "','" & Format(rstL("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstL("fecha_vencimiento"), "YYYY-mm-dd") & "','" & rstL("id_cliente") & "','" & rstL("ncliente") & "','" & rstL("documento") & "','" & rstL("total") * -1 & "','" & rstL("pago") * -1 & "','si','" & rstL("id_doc") & "','" & KEY_USUARIO & "','" & rstL("ruc") & "')"
                        CnBd.Execute (strCadena)
                        rstL.MoveNext
                    Next j
            End If
            
        End If
            
      
      
      
    End If
      
   


strCadena = "SELECT id_parametro,descripcion,fecha_emision,fecha_vencimiento,id_cliente,ncliente,documento,total,pago,saldo,'0',cobranza_dudosa,id_doc,ruc FROM view_reporte_gerencial_detallado WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "rpt_cuentas_cobrar_detalle", param, App.Path + "\Reportes\")

End Sub

Private Sub cmdInconsistencias_Click()
Dim in_monto As Single
Dim in_total As Single

strCadena = "SELECT v.id_venta,vv.documento,vv.total FROM venta_vendedor v,movimiento_venta vv WHERE v.id_venta=vv.id_venta and v.ruc=vv.ruc and   v.dni_save='" & KEY_USUARIO & "' and v.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
    in_total = Round(Format(rst("total"), "#,##0.00"), 2)
    in_monto = Round(get_total_comprobante(rst("id_venta")), 2)
        If in_total <> in_monto Then
            MsgBox "INCONSISTENCIA..." + rst("documento") + Chr(13) + Chr(13) + "TOTAL    :" + str(rst("total")) + Chr(13) + "DETALLE :" + str(in_monto)
        End If
        rst.MoveNext
   Next i
End If

End Sub
Private Function get_total_comprobante(ByVal in_venta As String) As Single

strCadena = "SELECT if(isnull(sum(total)),0,sum(total)) FROM movimiento_venta_detalle WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
get_total_comprobante = rstK(0)

End Function

Private Sub cmdLiquidacionLinea_Click()
        
        
        Call put_semanas

        strCadena = "call cursor_top_semana('" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        strCadena = "SELECT * FROM top_semana WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY id_semana ASC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
                strCadena = "call cursor_top_producto('" & rst("id_semana") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
         End If
          
   

   
   strCadena = "SELECT (can_sem1+can_sem2+can_sem3+can_sem4+can_sem5+can_sem6)/6,id_producto FROM view_top_rotacion WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      
      For i = 0 To rst.RecordCount - 1
          strCadena = "UPDATE producto SET stock_minimo='" & rst(0) & "' WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
          rst.MoveNext
          DoEvents
      Next i
   End If
   
   
   strCadena = "SELECT id_producto,nombre_prod,linea,stock,precio_venta,precio_compra,sem1,sem2,sem3,sem4,sem5,sem6,can_sem1,can_sem2,can_sem3," & _
   "can_sem4,can_sem5,can_sem6 FROM view_top_rotacion WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   
   Ans = ShowMultiReport(rst, "rpt_resumen_top", , App.Path + "\Reportes\")
   
   


End Sub
Private Sub put_semanas()
Dim fecha As Date



strCadena = "DELETE FROM top_semana WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



strCadena = "DELETE FROM top_semana_producto WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



fin_de_semana = FormatDateTime(CVDate(KEY_FECHA) - Weekday(KEY_FECHA) + 8, vbGeneralDate)
For i = 1 To 6
    
    fin_de_semana = DateAdd("d", -7, fin_de_semana)
    ini_de_semana = DateAdd("d", -6, fin_de_semana)
    no_sema = DatePart("ww", fin_de_semana, vbMonday, vbFirstFourDays)
    
    strCadena = "INSERT INTO top_semana(`semana`,`fecha_inicio`,`fecha_fin`,`dni_save`,`ruc`)VALUES " & _
    "('" & no_sema & "','" & Format(ini_de_semana, "YYYY-mm-dd") & "','" & Format(fin_de_semana, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
Next i



End Sub



Private Sub cmdmensual_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"



arr(0, 2) = Format(Me.DtpInicio_gerencial.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin_gerencial.Value, "dd-mm-YYYY")


param = arr()


If Me.OptReporteMensual.Value = True Then
    Call put_generar_resumen_mensual
    strCadena = "SELECT id,descripcion,monto_facturado,monto_cobrado,saldo_soles,saldo_dolares,numero_clientes,tc FROM reporte_cuentas_cobrar_resumen WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptResumen_credito_mensual", , App.Path + "\Reportes\")
End If

End Sub

Private Sub cmdRendimientoDetalle_Click()

        strCadena = "DELETE FROM venta_vendedor WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        strCadena = "call put_venta_vendedor_venta('" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "','" & DtcvendedorRendimiento.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        
         strCadena = "SELECT id_venta,id_vendedor,nombre_completo,fecha_emision,ncliente,documento,total,id_tipo FROM view_rendimiento_vendedor WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
                   
        
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "rpt_venta_vendedor_v2", , App.Path + "\Reportes\")
End Sub
Private Sub put_reporte()

Dim in_fecha_ini As Date
strCadena = "DELETE FROM venta_vendedor WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



strCadena = "SELECT DISTINCT id_vendedor FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   strCadena = "call put_ventas_comision('" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "','" & Format(DtpFinRendimiento.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
   CnBd.Execute (strCadena)
   
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "call put_venta_vendedor_venta_13('" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If

End Sub


Private Sub cmdreorterendimiento_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "fecha_ini"
arr(1, 1) = "fecha_fin"


arr(0, 2) = Format(Me.DtpInicioRendimiento.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFinRendimiento.Value, "dd-mm-YYYY")


param = arr()



If Me.OptReporteventas.Value = True Then
    Call ventas_vendedor_v2
    
    strCadena = "SELECT * FROM view_reporte_vendedor_linea WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptProduccion_vendedor", , App.Path + "\Reportes\")
    Exit Sub
End If

If opt_planilla_cobrador.Value = True Then
    
    If KEY_RUC = "20603698852" Or KEY_RUC = "20480431511" Then
        If Me.chk_cobrador.Value = 1 Then
            strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,id_cliente,ncliente,documento,total,dias,saldo,id_vendedor,vendedor FROM view_comprobantes_cobrar WHERE saldo>0 and  id_vendedor='" & Trim(DtcvendedorRendimiento.BoundText) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,id_cliente,ncliente,documento,total,dias,saldo,id_vendedor,vendedor FROM view_comprobantes_cobrar WHERE  fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        End If
        
        
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptCobranza_vendedor", param, App.Path + "\Reportes\")
        Exit Sub
    Else
    
    If Me.chk_cobrador.Value = 1 Then
    
        strCadena = "SELECT * FROM view_reporte_cobrador WHERE cobrador='" & Trim(DtcvendedorRendimiento.Text) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        'strCadena = "SELECT * FROM view_reporte_vendedor WHERE cobrador='" & Trim(DtcvendedorRendimiento.BoundText) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT * FROM view_reporte_cobrador WHERE fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    End If
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptCobranza_cobrador", , App.Path + "\Reportes\")
    End If
End If





If Me.opt_gasto_personal.Value = True Then
    
    If Me.chk_cobrador.Value = 1 Then
        If Me.chk_cuenta_contable.Value = 1 Then
            strCadena = "SELECT dni,nombre_completo,monto,fecha_emision,documento,cta_redondeo FROM view_gastos_contables WHERE cta_redondeo='" & Trim(Me.txtcuentacontable.Text) & "' and  dni='" & Trim(DtcvendedorRendimiento.Text) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT dni,nombre_completo,monto,fecha_emision,documento,cta_redondeo FROM view_gastos_contables WHERE dni='" & Trim(DtcvendedorRendimiento.Text) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        End If
    Else
           If Me.chk_cuenta_contable.Value = 1 Then
                strCadena = "SELECT dni,nombre_completo,monto,fecha_emision,documento,cta_redondeo FROM view_gastos_contables WHERE cta_redondeo='" & Trim(Me.txtcuentacontable.Text) & "' and fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
           Else
                strCadena = "SELECT dni,nombre_completo,monto,fecha_emision,documento,cta_redondeo FROM view_gastos_contables WHERE  fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
           End If
    End If
    
    
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptGastos_personal", param, App.Path + "\Reportes\")
    Exit Sub
    
End If




If opt_gastos_personal.Value = True Then
    
    If Me.chk_cobrador.Value = 1 Then
        
        strCadena = "SELECT `id_responsable`,`vendedor`,`fecha_emision`,`documento`,`observacion`,`total` ,`ruc` From view_gasto_personal_vii WHERE fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  id_responsable='" & DtcvendedorRendimiento.BoundText & "' and ruc='" & KEY_RUC & "'"
        
    Else
           strCadena = "SELECT `id_responsable`,`vendedor`,`fecha_emision`,`documento`,`observacion`,`total` ,`ruc` From view_gasto_personal_vii WHERE fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "'  and ruc='" & KEY_RUC & "'"
    End If
    
    
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptGastoPersonal", param, App.Path + "\Reportes\")
    
    
End If






'produccion por vendedor

If Me.OptReporteventas.Value = True Then
    
    If DateDiff("d", DtpInicioRendimiento.Value, Me.DtpFinRendimiento.Value) > 0 Then
    
        strCadena = "SELECT * FROM view_rendimiento_personal WHERE fecha_suscripcion>='" & Format(DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_suscripcion<='" & Format(DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
        
    Else
        strCadena = "SELECT * FROM view_rendimiento_personal WHERE   ruc='" & KEY_RUC & "'"
    End If
    
    Call ConfiguraRst(strCadena)
    Exit Sub
End If


If Me.opt_produccion.Value = True Then
   
   strCadena = "SELECT `dni`,`nombre_completo`,`direccion`,`celular`,`id_vendedor`,`vendedor`,`fecha_suscripcion`,`id_plan`,`plan`,`monto`,`pago_mensual`,`pago_3meses`,`pago_6meses`,`pago_anual`,`estado`,`medio_contacto`,`prioridad`,`observacion` FROM view_rendimiento_personal WHERE fecha_suscripcion>='" & Format(DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and fecha_suscripcion<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptRendimientoPersonal", param, App.Path + "\Reportes\")
   Exit Sub
End If





End Sub
Private Sub planilla_cobranza()


End Sub


Private Sub ventas_vendedor()
Dim codigo_linea() As String
Dim descripcion_linea() As String
Dim monto_linea() As Double

strCadena = "DELETE  FROM reporte_vendedor_linea WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM view_produccion_vendedor_linea WHERE  dni_save='" & KEY_USUARIO & "' and   ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_acumulado_nota = 0
   in_ventas = 0
   in_comision = 0
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea"
       Call ConfiguraRstT(strCadena)
       If rstT.RecordCount > 0 Then
          rstT.MoveFirst
          ReDim codigo_linea(rstT.RecordCount - 1)
          ReDim descripcion_linea(rstT.RecordCount - 1)
          ReDim monto_linea(rstT.RecordCount - 1)
          For j = 0 To rstT.RecordCount - 1
              codigo_linea(j) = rstT("id_linea")
              descripcion_linea(j) = rstT("descripcion")
              
              strCadena = "SELECT sum(d.`total`) From movimiento_venta v,movimiento_venta_detalle d,producto p Where v.`id_venta`=d.`id_venta` and  " & _
              " d.`id_producto`=p.`id_producto` and d.`ruc`=p.`ruc` and v.`ruc`=d.`ruc` and v.`id_vendedor`='" & rst("id_vendedor") & "' and v.`fecha_emision`>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and " & _
              " v.`fecha_emision`<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and p.`id_linea`='" & rstT("id_linea") & "' and " & _
              " v.ruc='" & KEY_RUC & "' and v.`id_doc` IN('0001','0003') and v.anulado='no'"
              Call ConfiguraRstP(strCadena)
              If IsNull(rstP(0)) = False Then
                monto_linea(j) = rstP(0)
                in_ventas = rstP(0)
              Else
                monto_linea(j) = 0
                in_ventas = 0
              End If
              in_comision = in_ventas + in_comision
              in_nota = 0
              strCadena = "SELECT sum(d.`total`) From movimiento_venta v,movimiento_venta_detalle d,Producto p,movimiento_venta vv Where v.`id_venta`=d.`id_venta` and  " & _
              " d.`id_producto`=p.`id_producto` and d.`ruc`=p.`ruc` and v.`ruc`=d.`ruc` and v.`id_vendedor`='" & rst("id_vendedor") & "'  and " & _
              " v.`fecha_emision`>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and v.`fecha_emision`<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "' and p.`id_linea`='" & rstT("id_linea") & "' and " & _
              " v.ruc='" & KEY_RUC & "' and v.`id_doc` IN('0007') and v.id_comprobante=vv.id_venta and v.ruc=vv.ruc and vv.fecha_emision>='" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "' and vv.fecha_emision<='" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "'"
              Call ConfiguraRstP(strCadena)
             If IsNull(rstP(0)) = False Then
                in_nota = rstP(0)
              Else
                in_nota = 0
              End If
              monto_linea(j) = monto_linea(j) - in_nota
              in_acumulado_nota = in_acumulado_nota + in_nota
              rstT.MoveNext
          Next j
       End If
              strCadena = "INSERT INTO reporte_vendedor_linea(`fecha_inicio`,`fecha_fin`,`dni_vendedor`,`id_linea1`,`id_linea2`,`id_linea3`,`id_linea4`,`id_linea5`,`id_linea6`,`linea1`,`linea2`,`linea3`,`linea4`,`linea5`,`linea6`,`monto_linea1`," & _
              "`monto_linea2`,`monto_linea3`,`monto_linea4`,`monto_linea5`,`monto_linea6`,`dni_save`,`ruc`)VALUES('" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "', " & _
              "'" & rst("id_vendedor") & "','" & codigo_linea(0) & "','" & codigo_linea(1) & "','" & codigo_linea(2) & "','" & codigo_linea(3) & "','" & codigo_linea(4) & "','" & codigo_linea(5) & "','" & descripcion_linea(0) & "','" & descripcion_linea(1) & "'," & _
              "'" & descripcion_linea(2) & "','" & descripcion_linea(3) & "','" & descripcion_linea(4) & "','" & descripcion_linea(5) & "','" & monto_linea(0) & "','" & monto_linea(1) & "','" & monto_linea(2) & "','" & monto_linea(3) & "','" & monto_linea(4) & "','" & monto_linea(5) & "'," & _
              "'" & KEY_USUARIO & "','" & KEY_RUC & "')"
              CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub

Private Sub ventas_vendedor_v2()
Dim codigo_linea() As String
Dim descripcion_linea() As String
Dim monto_linea() As Double



Call put_reporte
strCadena = "DELETE  FROM reporte_vendedor_linea WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM view_produccion_vendedor_linea WHERE  dni_save='" & KEY_USUARIO & "' and   ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_acumulado_nota = 0
   in_ventas = 0
   in_comision = 0
   
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rst.RecordCount
   
   For i = 0 To rst.RecordCount - 1
              
       strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea"
       Call ConfiguraRstT(strCadena)
       If rstT.RecordCount > 0 Then
          rstT.MoveFirst
          ReDim codigo_linea(rstT.RecordCount - 1)
          ReDim descripcion_linea(rstT.RecordCount - 1)
          ReDim monto_linea(rstT.RecordCount - 1)
          For j = 0 To rstT.RecordCount - 1
              codigo_linea(j) = rstT("id_linea")
              descripcion_linea(j) = rstT("descripcion")
              
              strCadena = "SELECT if (isnull(sum(total)),0,sum(total))  FROM view_vendedor_linea WHERE id_linea='" & rstT("id_linea") & "' and  id_vendedor='" & rst("id_vendedor") & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
              Call ConfiguraRstL(strCadena)
              monto_linea(j) = rstL(0)
              DoEvents
              rstT.MoveNext
          Next j
       End If
              strCadena = "INSERT INTO reporte_vendedor_linea(`fecha_inicio`,`fecha_fin`,`dni_vendedor`,`id_linea1`,`id_linea2`,`id_linea3`,`id_linea4`,`id_linea5`,`id_linea6`,`linea1`,`linea2`,`linea3`,`linea4`,`linea5`,`linea6`,`monto_linea1`," & _
              "`monto_linea2`,`monto_linea3`,`monto_linea4`,`monto_linea5`,`monto_linea6`,`dni_save`,`ruc`)VALUES('" & Format(Me.DtpInicioRendimiento.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFinRendimiento.Value, "YYYY-mm-dd") & "', " & _
              "'" & rst("id_vendedor") & "','" & codigo_linea(0) & "','" & codigo_linea(1) & "','" & codigo_linea(2) & "','" & codigo_linea(3) & "','" & codigo_linea(4) & "','" & codigo_linea(5) & "','" & descripcion_linea(0) & "','" & descripcion_linea(1) & "'," & _
              "'" & descripcion_linea(2) & "','" & descripcion_linea(3) & "','" & descripcion_linea(4) & "','" & descripcion_linea(5) & "','" & monto_linea(0) & "','" & monto_linea(1) & "','" & monto_linea(2) & "','" & monto_linea(3) & "','" & monto_linea(4) & "','" & monto_linea(5) & "'," & _
              "'" & KEY_USUARIO & "','" & KEY_RUC & "')"
              CnBd.Execute (strCadena)
       
       
       rst.MoveNext
       DoEvents
       Me.prg_avance.Value = i
       DoEvents
    Next i
End If
End Sub

Private Sub cmdReport_Click()
Call put_semanas

        strCadena = "call cursor_top_semana_margen('" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        strCadena = "SELECT * FROM top_semana WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY id_semana ASC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
                strCadena = "call cursor_top_producto('" & rst("id_semana") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
         End If
          
   

   
   strCadena = "SELECT (can_sem1+can_sem2+can_sem3+can_sem4+can_sem5+can_sem6)/6,id_producto FROM view_top_rotacion WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
          strCadena = "UPDATE producto SET stock_minimo='" & rst(0) & "' WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
          rst.MoveNext
          
      Next i
   End If




strCadena = "DELETE FROM producto_quiebre WHERE  ruc='" & KEY_RUC & " '"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM view_quiebre_producto WHERE ruc='" & KEY_RUC & "' LIMIT 400"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("stock_minimo") <= rst("margen") Then
       strCadena = "INSERT INTO producto_quiebre(`id_producto`,`stock_minimo`,`margen`,`dni_save`,`ruc`)VALUES" & _
       "('" & rst("id_producto") & "','" & rst("stock_minimo") & "','" & D & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       End If
       rst.MoveNext
   Next i
End If



If Me.chk_lineaquiebre.Value = 1 Then
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,stock_minimo,margen FROM view_quiebre_reporte WHERE id_linea='" & Me.DtcLineaQuiebre.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' "
Else
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,stock_minimo,margen FROM view_quiebre_reporte WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' "
End If

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptStockMinimo", , App.Path + "\Reportes\")
    
End Sub

Private Sub cmdReporte_comparativo_Click()
If Me.Opt_comparativo_diario.Value = True Then
   in_criterio = ""
   If Me.chk_linea_c.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca_c.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca_c.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo_c.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo_c.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_proveedor_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor_c.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor_c.BoundText & "'"
      End If
   
   End If
   
   
   
   If Me.chk_sucursal_c.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio_c.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin_c.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen_c.Text & "',id_producto,detalle,sum(cantidad),linea,marca,id_vendedor,documento,precio,precio_compra,sum(total),fecha_emision FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio_c.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_c.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen_c.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY fecha_emision,id_alm  ORDER BY 14 ASC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,sum(cantidad),linea,marca,id_vendedor,documento,precio,precio_compra,sum(total),fecha_emision FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio_c.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_c.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY fecha_emision ORDER BY 14 ASC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptVenta_valorizado", , App.Path + "\Reportes\")
   Exit Sub

End If

If Me.Opt_comparativo_mensual.Value = True Then
   in_criterio = ""
   If Me.chk_linea_c.Value = 1 Then
      in_criterio = " and id_linea='" & Me.DtcLinea.BoundText & "'"
  
   End If
   If Me.chk_marca_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and  id_marca='" & Me.DtcMarca_c.BoundText & "'" & in_criterio
      Else
        in_criterio = "and id_marca='" & Me.DtcMarca_c.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_modelo_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_sublinea='" & Me.DtcModelo_c.BoundText & "'" & in_criterio
      Else
        in_criterio = " and id_sublinea='" & Me.DtcModelo_c.BoundText & "'"
      End If
   
   End If
   
   If Me.chk_proveedor_c.Value = 1 Then
      If in_criterio <> "" Then
        in_criterio = "and id_proveedor='" & Me.DtcProveedor_c.BoundText & "' " & in_criterio
      Else
        in_criterio = "and id_proveedor='" & Me.DtcProveedor_c.BoundText & "'"
      End If
   
   End If
   
   
   
   If Me.chk_sucursal_c.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio_c.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin_c.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen_c.Text & "',id_producto,detalle,sum(cantidad),linea,marca,id_vendedor,documento,precio,precio_compra,sum(total),fecha_emision FROM view_producto_rotacion WHERE  fecha_emision>='" & Format(Me.DtpInicio_c.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_c.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen_c.BoundText & "' and ruc='" & KEY_RUC & "' " & in_criterio & " GROUP BY month(fecha_emision),id_alm  ORDER BY 14 ASC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "','TODAS LAS SUCURSALES',id_producto,detalle,sum(cantidad),linea,marca,id_vendedor,documento,precio,precio_compra,sum(total),fecha_emision FROM view_producto_rotacion WHERE fecha_emision>='" & Format(Me.DtpInicio_c.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_c.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' " & in_criterio & "GROUP BY month(fecha_emision) ORDER BY 14 ASC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptVenta_valorizado", , App.Path + "\Reportes\")
   Exit Sub

End If


End Sub

Private Sub cmdTransacciones_Click()


strCadena = "CALL cursor_transacciones('" & Format(Me.DtpInicio_transacciones.Value, "YYYY-mm-dd") & "','" & Format(Me.Dtpfin_transacciones.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT `id_hora`,CONCAT(`inicio`,'-',`hora_fin`) as periodo,`atenciones`,`ventas` FROM horas WHERE ruc='" & KEY_RUC & "' ORDER BY CAST(inicio as TIME)"

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_atenciones", , App.Path + "\Reportes\")


End Sub

Private Sub cmdUpdateCosto_Click()
Dim in_costo As Double
If MsgBox("Desea Actualizar los costos", vbYesNo + vbQuestion) = vbYes Then
    strCadena = "CALL ADM_reportes_generales('6','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','','" & Me.DtcAlmacen.BoundText & "','','','','','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
            strCadena = "CALL ADM_reportes_generales('7','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_producto") & "','" & Me.DtcAlmacen.BoundText & "','','','','','" & KEY_RUC & "')"
            Call ConfiguraRstIN(strCadena)
            
            If Round(rstIN(0), 6) <> rst("precio_costo") Then
                strCadena = "UPDATE movimiento_venta_detalle SET precio_costo='" & rstIN(0) & "' WHERE  id_detalle_venta='" & rst("id_detalle_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
            End If
            rst.MoveNext
            
            DoEvents
            Me.cmdUpdateCosto.Caption = str(i) & Space(2) & str(rst.RecordCount)
       Next i
    End If
End If

End Sub

Private Sub DtcLinea_Change()
strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcModelo)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
  Me.DtpFin.Value = KEY_FECHA
  Me.DtpInicio.Value = KEY_FECHA
  Me.DtpInicioRendimiento.Value = KEY_FECHA
  Me.DtpFinRendimiento.Value = KEY_FECHA
  Me.DtpFechaReporte_gerencial.Value = KEY_FECHA
  
  Me.DtpFin_gerencial.Value = KEY_FECHA
  Me.DtpInicio_gerencial.Value = KEY_FECHA
  Me.DtpInicio_transacciones.Value = KEY_FECHA
  Me.Dtpfin_transacciones.Value = KEY_FECHA
  Me.DTPDiferidaIni.Value = KEY_FECHA
  Me.DTPDiferidaFin.Value = KEY_FECHA
  
  
  
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcLinea)
  
  
  
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinearendimiento)
  
  
  
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea_c)
  
  strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcModelo)
  
  strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM linea_modelo WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtcModelo.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcModelo_ii)
  
  
  strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcModelo_c)
  
  strCadena = "SELECT id_marca as Codigo,descripcion as Descripcion FROM marca WHERE  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
  
  strCadena = "SELECT id_marca as Codigo,descripcion as Descripcion FROM marca WHERE  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarcaRendimiento)
  
  strCadena = "SELECT id_marca as Codigo,descripcion as Descripcion FROM marca WHERE  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca_c)
  
  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  
  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcvendedorRendimiento)
  
  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtpVendedor_c)
  

  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_proveedor='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor_c)
  
  
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_cliente='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCliente)
  
  
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcSucursalRendimiento)
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen_c)
  
  
  

  
  
  
  
  Me.DtcAlmacen.BoundText = KEY_ALM
  
  
  
  If KEY_RUC = "20480516771" Then
     Me.opt_descripcion.Value = True
  End If
  
  



  
  
End Sub






Private Sub Opt_reporte103_Click()
Me.cmdDetallado.Visible = True
End Sub

Private Sub opt_ventas_vendedor_Click()
Me.cmdDetallado.Visible = False
End Sub

Private Sub TxtBuscarCliente_Change()
 strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.TxtBuscarCliente.Text) & "%' and  id_cliente='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCliente)
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub txtBusquedaCliente_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBusquedaCliente.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcBusquedaCliente)
End Sub

Private Sub txtBusquedaproveedor_Change()
 strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBusquedaCliente.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
End Sub

Private Sub txtCuentaContable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPlanContableCuentas.Show
End If
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

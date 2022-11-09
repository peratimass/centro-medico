VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmKardexdeProductos 
   BorderStyle     =   0  'None
   Caption         =   "Kardex de Productos"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_periodo 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00800000&
      Height          =   310
      Left            =   1920
      TabIndex        =   47
      Top             =   840
      Width           =   375
   End
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   315
      Left            =   2400
      TabIndex        =   46
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.CommandButton cmdCorregirCosto 
      Caption         =   "CORREGIR INGRESOS"
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
      Left            =   7560
      TabIndex        =   40
      Top             =   840
      Width           =   2295
   End
   Begin VB.CheckBox chk_costo_promedio 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "COSTO PROMEDIO"
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
      Height          =   240
      Left            =   17760
      TabIndex        =   38
      Top             =   0
      Width           =   1935
   End
   Begin VB.CheckBox chk_all 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "ALL"
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
      Left            =   4680
      TabIndex        =   35
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame frmCuadrar 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   850
      Left            =   135
      TabIndex        =   30
      Top             =   1215
      Width           =   4455
      Begin VitekeySoft.ChameleonBtn cmdCuadrarSalidas 
         Height          =   350
         Left            =   120
         TabIndex        =   31
         Top             =   75
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "CUADRAR SALIDAS"
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
         MICON           =   "FrmKardexProducto.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCuadrarIngresos 
         Height          =   350
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "CUADRAR INGRESOS"
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
         MICON           =   "FrmKardexProducto.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsaldoInicial 
         Height          =   350
         Left            =   2640
         TabIndex        =   36
         Top             =   75
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "UPDATE SALDO INICIAL"
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
         MICON           =   "FrmKardexProducto.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreportevalorizado 
         Height          =   350
         Left            =   2640
         TabIndex        =   44
         Top             =   480
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "REPORTE VALORIZADO"
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
         MICON           =   "FrmKardexProducto.frx":0054
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
   Begin VB.CheckBox chk_sinprocesar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "SIN PROCESAR"
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
      Height          =   230
      Left            =   4920
      TabIndex        =   26
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CORREGIR SALDO STOCK"
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
      Left            =   9960
      TabIndex        =   24
      Top             =   840
      Width           =   2895
   End
   Begin VB.CheckBox chk_ventas_realizadas 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VENTAS REALIZADAS"
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
      Height          =   240
      Left            =   15720
      TabIndex        =   19
      Top             =   620
      Width           =   1935
   End
   Begin VB.CheckBox chk_kardex_contable 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "KARDEX CONTABLE"
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
      Height          =   240
      Left            =   15720
      TabIndex        =   18
      Top             =   340
      Width           =   1935
   End
   Begin VB.CheckBox chk_kardex_fisico 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "KARDEX FISICO"
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
      Height          =   240
      Left            =   15720
      TabIndex        =   17
      Top             =   80
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VitekeySoft.ChameleonBtn cmdactivar 
      Height          =   350
      Left            =   5400
      TabIndex        =   14
      Top             =   120
      Width           =   550
      _ExtentX        =   979
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmKardexProducto.frx":0070
      PICN            =   "FrmKardexProducto.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkBuscarfechas 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "BUSCAR X FECHAS :"
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
      Height          =   230
      Left            =   4920
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   820
      Left            =   6720
      TabIndex        =   9
      Top             =   1140
      Visible         =   0   'False
      Width           =   13215
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   225
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   63504385
         CurrentDate     =   41291
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   225
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   63504385
         CurrentDate     =   41291
      End
      Begin VitekeySoft.ChameleonBtn cmdKardexGeneral 
         Height          =   280
         Left            =   4995
         TabIndex        =   21
         Top             =   165
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "KARDEX  GENERAL"
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
         MICON           =   "FrmKardexProducto.frx":25B7
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
         Height          =   250
         Left            =   3160
         TabIndex        =   22
         Top             =   200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmd_inconsistencia 
         Height          =   585
         Left            =   9120
         TabIndex        =   25
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1032
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
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmKardexProducto.frx":25D3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmreportedetallado 
         Height          =   280
         Left            =   10800
         TabIndex        =   27
         Top             =   465
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "DOC. INGRESOS KARDEX"
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
         MICON           =   "FrmKardexProducto.frx":25EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdPeriodoComrpobantes 
         Height          =   280
         Left            =   10800
         TabIndex        =   28
         Top             =   160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "DOC. SALIDAS KARDEX"
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
         MICON           =   "FrmKardexProducto.frx":260B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdKardexResumido 
         Height          =   280
         Left            =   4995
         TabIndex        =   29
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "KARDEX RESUMEN"
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
         MICON           =   "FrmKardexProducto.frx":2627
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdInventarioValorizado 
         Height          =   285
         Left            =   6720
         TabIndex        =   33
         Top             =   165
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "F.13.1  INV.PERMAN.VAL"
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
         MICON           =   "FrmKardexProducto.frx":2643
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreportekardex 
         Height          =   285
         Left            =   6720
         TabIndex        =   41
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "REPORTE KARDEX "
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
         MICON           =   "FrmKardexProducto.frx":265F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
         Height          =   285
         Left            =   3160
         TabIndex        =   45
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "STOCK A LA FECHA"
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
         MICON           =   "FrmKardexProducto.frx":267B
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
         Caption         =   "AL"
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
         Left            =   1485
         TabIndex        =   12
         Top             =   240
         Width           =   195
      End
   End
   Begin VB.TextBox TxtBusquedarapida 
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
      Left            =   7560
      TabIndex        =   1
      Top             =   530
      Width           =   5295
   End
   Begin VB.TextBox TxtcodigoProd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12960
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
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
   Begin MSDataListLib.DataCombo DtcProducto 
      Height          =   330
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   6735
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   11880
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
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   255
      Left            =   19800
      TabIndex        =   15
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmKardexProducto.frx":2697
      PICN            =   "FrmKardexProducto.frx":26B3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdKardexContable 
      Height          =   375
      Left            =   12960
      TabIndex        =   16
      Top             =   525
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "  REPORTE KARDEX"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmKardexProducto.frx":5567
      PICN            =   "FrmKardexProducto.frx":5583
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   345
      Left            =   210
      TabIndex        =   20
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "UPDATE KARDEX"
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
      MICON           =   "FrmKardexProducto.frx":7B68
      PICN            =   "FrmKardexProducto.frx":7B84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdUpdateKardexTotal 
      Height          =   345
      Left            =   1920
      TabIndex        =   23
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "UPDATE A TODO"
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
      MICON           =   "FrmKardexProducto.frx":A0AF
      PICN            =   "FrmKardexProducto.frx":A0CB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCostopromedio 
      Height          =   375
      Left            =   17760
      TabIndex        =   37
      Top             =   300
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "SALDO STOCK            "
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
      MICON           =   "FrmKardexProducto.frx":C5F6
      PICN            =   "FrmKardexProducto.frx":C612
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   17760
      TabIndex        =   39
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "COSTO PROMEDIO"
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
      MICON           =   "FrmKardexProducto.frx":EBF7
      PICN            =   "FrmKardexProducto.frx":EC13
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
      Height          =   195
      Left            =   210
      TabIndex        =   42
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar progress_costo 
      Height          =   195
      Left            =   15720
      TabIndex        =   43
      Top             =   900
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMATO 13.1 REGISTRO DE INVENTARIO PERMANENTE VALORIZADO"
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
      Height          =   225
      Left            =   105
      TabIndex        =   34
      Top             =   2040
      Width           =   5715
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S  A  L  D  O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   16905
      TabIndex        =   8
      Top             =   2115
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S  A  L  I  D  A  S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   12495
      TabIndex        =   7
      Top             =   2115
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I N G R E S O S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   8340
      TabIndex        =   6
      Top             =   2115
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSQUEDA RAPIDA:"
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
      Left            =   6075
      TabIndex        =   5
      Top             =   600
      Width           =   1365
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11200
      Top             =   2100
      Width           =   3975
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15280
      Top             =   2100
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7080
      Top             =   2100
      Width           =   4040
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   4800
      Top             =   1125
      Width           =   15255
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmKardexdeProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Function get_costo_unit(ByVal in_producto As String, ByVal in_alm As String)
strCadena = "SELECT precio_compra FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    get_costo_unit = rstZ("precio_compra")

End If
End Function
Private Function get_precio_compra_unit(ByVal in_producto As String)
strCadena = "SELECT c_unitario FROM view_compra_detalle WHERE c_unitario>0.01 and  id_producto='" & in_producto & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY valor_venta DESC LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_precio_compra_unit = rstK(0)
Else
    strCadena = "SELECT precio FROM movimiento_venta_detalle WHERE  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1 "
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        get_precio_compra_unit = rstK(0) - rstK(0) * 20 / 100
    End If
End If

End Function
Private Sub ChameleonBtn1_Click()

Dim in_costo_sta As Double
Me.DtcAlmacen.Enabled = True
Me.DtcProducto.Enabled = True
Me.TxtcodigoProd.Enabled = True

If KEY_RUC = "20487376338" Then
    strCadena = "SELECT * FROM kardex WHERE id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       If rstA("costo_unitario") <= 0.1 Then
              in_costo_sta = get_costo_unit(Trim(Me.TxtcodigoProd.Text), Me.DtcAlmacen.BoundText)
              strCadena = "UPDATE kardex SET costo_unitario='" & Val(in_costo_sta) & "',costo_promedio='" & Val(in_costo_sta) & "' WHERE id_kardex='" & rstA("id_kardex") & "' and   id_producto='" & Trim(Me.TxtcodigoProd.Text) & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
              CnBd.Execute (strCadena)
       End If
    End If
End If

If Trim(Me.TxtcodigoProd.Text) <> "" Then
    strCadena = "SELECT * FROM view_producto WHERE   id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and  id_tipo='01'  and id_alm='" & Me.DtcAlmacen.BoundText & "'    and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_producto WHERE     id_alm='" & Me.DtcAlmacen.BoundText & "'    and ruc='" & KEY_RUC & "' ORDER BY id_producto DESC"
End If

Call ConfiguraRstC(strCadena)
If rstc.RecordCount > 0 Then
   rstc.MoveFirst
   Me.progress_costo.Min = 0
   Me.progress_costo.Max = rstc.RecordCount + 1
   For i = 0 To rstc.RecordCount - 1
    
        Me.DtcAlmacen.BoundText = rstc("id_alm")
        Me.TxtcodigoProd.Text = Trim(rstc("id_producto"))
        'Call put_update_saldo_stock_contable(rstc("id_producto"), rstc("id_alm"))
        
        'Call put_update_saldo_stock_fisico(rstc("id_producto"), rstc("id_alm"))
        
        
        Call put_update_costo_promedio_fisico(rstc("id_producto"), rstc("id_alm"))
        
        'Call put_update_costo_promedio_contable(rstc("id_producto"), rstc("id_alm"))
        
          DoEvents
       ' Call load_kardex
         
        DoEvents
        progress_costo.Value = i
        rstc.MoveNext
        DoEvents
   Next i
   MsgBox "listo"
End If



End Sub

Private Sub chk_kardex_contable_Click()
If Me.chk_kardex_contable.Value = 1 Then
   Me.chk_kardex_fisico.Value = 0
Else
   Me.chk_kardex_fisico.Value = 1
End If
End Sub

Private Sub chk_kardex_fisico_Click()
If Me.chk_kardex_fisico.Value = 1 Then
   Me.chk_kardex_contable.Value = 0
Else
   Me.chk_kardex_contable.Value = 1
End If
End Sub

Private Sub chk_periodo_Click()

If Me.chk_periodo.Value = 1 Then
   
   strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcPeriodo)
   Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
   Me.DtcPeriodo.Visible = True
Else
    Me.DtcPeriodo.Visible = False
End If


End Sub

Private Sub chkBuscarfechas_Click()
If Me.chkBuscarfechas.Value = 1 Then
    Me.Frame1.Visible = True
    Me.DtpDesde.Value = KEY_FECHA
    Me.DtpHasta.Value = KEY_FECHA
Else
    Me.Frame1.Visible = False
End If
End Sub

Private Sub cmd_inconsistencia_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String



strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto DESC "
ConfiguraRst (strCadena)
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 2 Then
            
            strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,stock,precio_venta,precio_compra,habilitado,ruc)VALUES " & _
            "('00002','" & rst("id_producto") & "','0','" & rstL("precio_venta") & "','" & rstL("precio_compra") & "','si','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,stock,precio_venta,precio_compra,habilitado,ruc)VALUES " & _
            "('00003','" & rst("id_producto") & "','0','" & rstL("precio_venta") & "','" & rstL("precio_compra") & "','si','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
      End If
            rst.MoveNext
   Next i
End If


Exit Sub




arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")

param = arr()

strCadena = "DELETE FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' ORDER BY k.id_producto"

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rst.RecordCount - 1
   
   strCadena = "call put_crear_kardex_temporal('" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "');"
   CnBd.Execute (strCadena)
   
   
   For i = 0 To rst.RecordCount - 1
        strCadena = "CALL procedure_kardex_inconsistencia('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Me.prg_avance.Value = i
        Me.cmdKardexGeneral.Caption = (i + 1) & Space(2) & "-" & Space(1) & rst.RecordCount
        DoEvents
   Next i

End If
 strCadena = "SELECT id_producto,producto,cantidad_inicial,saldo_inicial,cantidad_ingreso,saldo_ingreso,cantidad_salida,saldo_salida,cantidad_final,saldo_final FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)


   




Ans = ShowMultiReport(rst, "RptKardexValorizado", param, App.Path + "\Reportes\")



End Sub

Private Sub cmdactivar_Click()
Me.DtcAlmacen.Enabled = True
Me.DtcProducto.Enabled = True
Me.TxtcodigoProd.Enabled = True

Call Resalta(Me.TxtcodigoProd)
End Sub

Private Sub cmdCerrar_Click()
Call cerrar_form
End Sub
Public Sub cerrar_form()
Unload Me
End Sub
Private Sub cmdcomprobantesIngreso_Click()
End Sub

Private Sub cmdCorregirCosto_Click()


strCadena = "SELECT * FROM orden_compra WHERE id_orden>6686 and  monto_flete>0 and  id_estado<>'3' and   id_doc='0414' and fecha_solicitud>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_solicitud<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY id_orden DESC "
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    
    For i = 0 To rstK.RecordCount
         Call put_flete_orden(rstK("id_orden"), rstK("monto_flete"), rstK("total"))
         DoEvents
         Call actualizar_kardex(rstK("id_orden"))
         
         
         
         
         If i = 0 Then
         If MsgBox("Esta correcto el calculo, para Proceder", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
         End If
         DoEvents
         rstK.MoveNext
         Me.cmdCorregirCosto.Caption = str(i) & Space(5) & str(rstK.RecordCount) & Space(2) & rstK("id_orden")
         DoEvents
    Next i
    
    'aqui tenemos que hacer distinct y recorrer todos los productos
End If
MsgBox "FINALIZADO CON EXITO"

End Sub
Private Sub put_flete_orden(ByVal in_recepcion As String, ByVal in_flete As Single, ByVal in_valor_venta As Double)
Dim in_monto_gasto As Single
strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & in_recepcion & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_monto_gasto = in_flete
   

   
   For i = 0 To rst.RecordCount - 1
        If in_monto_gasto > 0 Then
            in_monto_parcial = rst("precio") / (in_valor_venta) * in_monto_gasto
            in_monto_porcentaje = in_monto_parcial * 100 / in_monto_gasto
       Else
           in_monto_parcial = 0
           in_monto_porcentaje = 0
        
        End If
        
        
        strCadena = "UPDATE orden_compra_detalle SET incremento_neto='" & in_monto_parcial * (1 + KEY_IGV) & "'  WHERE id_detalle='" & rst("id_detalle") & "'"
        CnBd.Execute (strCadena)
        DoEvents
        rst.MoveNext
        
   Next i
   End If
End Sub

Private Sub actualizar_kardex(ByVal in_recepcion As String)
    Dim in_costo_igv As Single
    Dim in_afecto_igv As String
    Dim in_moneda As String
    Dim in_factor As Single
    
   'OBTENGO LA ORDEN DE COMPRA
   strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & Val(in_recepcion) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstA(strCadena)
   If rstA.RecordCount > 0 Then
        'ACTUALIZAR KARDEX CON GUIA
        
        
        strCadena = "SELECT serie,numero FROM movimiento_compra WHERE id_compra='" & rstA("id_compra") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            in_serie_compra = rstT("serie")
            in_numero_compra = rstT("numero")
        Else
            in_serie_compra = ""
            in_numero_compra = ""
            
        End If
   
        
        
        in_moneda = rstA("id_moneda")
        in_afecto_igv = rstA("afecto_igv")
        
        
        in_compra_flete = rstA("id_factura_flete")
        in_serie_guia = rstA("guia_serie")
        in_numero_guia = rstA("guia_numero")
        in_compra = rstA("id_compra")
        
         
                
                
                
        
     
            strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_recepcion) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstIN(strCadena)
            If rstIN.RecordCount > 0 Then
               rstIN.MoveFirst
               If in_moneda = "00001" Then
                   in_factor = 1
               Else
                   in_factor = KEY_CAMBIO
               End If
               For i = 0 To rstIN.RecordCount - 1
                    Me.TxtcodigoProd.Text = rstIN("id_producto")
                    in_monto_neto = (rstIN("precio") * in_factor + rstIN("incremento_neto")) * (1 + KEY_IGV)
                    
                    
                    If in_afecto_igv = "si" Then
                        in_costo_igv = rstIN("precio") * in_factor + rstIN("precio") * KEY_IGV * in_factor + rstIN("incremento_neto") * (1 + KEY_IGV)
                    Else
                        in_costo_igv = in_monto_neto
                    End If
                    
                     Me.TxtcodigoProd.Text = rstIN("id_producto")
                    
                   ' If Trim(in_serie_guia) = "" And Trim(in_numero_guia) = "" Then
                   '     strCadena = "call put_kardex_stock('04','" & Format(rstA("fecha_solicitud"), "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0001','" & Trim(in_serie_compra) & "','" & Trim(in_numero_compra) & "','" & rstA("id_proveedor") & "','" & rstIN("id_producto") & "','" & rstIN("cantidad") & "','" & in_costo_igv & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                   ' Else
                   '     strCadena = "call put_kardex_stock('04','" & Format(rstA("fecha_solicitud"), "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0009','" & Trim(in_serie_guia) & "','" & Trim(in_numero_guia) & "','" & Trim(rstA("id_proveedor")) & "','" & rstIN("id_producto") & "','" & rstIN("cantidad") & "','" & in_costo_igv & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                   ' End If
                    
                   ' CnBd.Execute (strCadena)
                   
                    
                   
                        Call update_kardex_VARGAS(Trim(rstIN("id_producto")))
                    
                    
                    rstIN.MoveNext
                    DoEvents
               Next i
            End If
        End If
   ' End If
End Sub

Private Sub cmdCostopromedio_Click()



Call put_verificar_kardex_producto(Trim(Me.TxtcodigoProd.Text))


End Sub

Private Sub put_saldo_stock()

strCadena = "SELECT * FROM moviimiento_venta "

End Sub










Private Sub put_update_costo_ingreso()
Dim in_costo As Double

strCadena = "SELECT * FROM view_producto WHERE id_alm='00001' and  ruc='" & KEY_RUC & "' and precio_venta=0"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_alm='" & rst("id_alm") & "' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY id_compra DESC LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            strCadena = "UPDATE almacen_producto SET precio_venta='" & rstL("p_venta") & "' WHERE   id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        rst.MoveNext
   Next i
End If

Exit Sub

strCadena = "SELECT * FROM view_compra_detalle WHERE valor_venta>0 and  cantidad>0 and  ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        If rst("id_moneda") = "00001" Then
            in_costo = (rst("valor_venta") + rst("incremento_neto_gasto")) / rst("cantidad")
        Else
            in_costo = (rst("valor_venta") * rst("tc") + rst("incremento_neto_gasto")) / rst("cantidad")
        End If
        strCadena = "UPDATE kardex SET costo_unitario='" & in_costo & "' WHERE cantidad_real>0 and  id_movimiento='" & rst("id_compra") & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        DoEvents
    Next i
End If
Exit Sub

strCadena = "SELECT * FROM view_compra_detalle WHERE valor_venta=0 and  id_doc='0089' and  ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        If rst("id_moneda") = "00001" Then
            in_costo = (rst("valor_venta") + rst("incremento_neto_gasto")) / rst("cantidad")
        Else
            in_costo = (rst("valor_venta") * rst("tc") + rst("incremento_neto_gasto")) / rst("cantidad")
        End If
        strCadena = "UPDATE kardex SET costo_unitario='" & in_costo & "' WHERE cantidad_real>0 and  id_movimiento='" & rst("id_compra") & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
        DoEvents
    Next i
End If


End Sub
Private Sub put_update_costos()

strCadena = "SELECT * FROM view_producto WHERE id_alm<>'00001' and  precio_compra*1.18>=precio_venta and  ruc='" & KEY_RUC & "' ORDER BY id_alm,id_producto ASC"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   For i = 0 To rstA.RecordCount - 1
       strCadena = "SELECT * FROM kardex WHERE     id_alm='" & rstA("id_alm") & "' and id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          rst.MoveFirst
          For j = 0 To rst.RecordCount - 1
                
                'obtengo el ultimo precio anterior a ese en almacen principal
                 strCadena = "SELECT costo_promedio,id_kardex,cantidad FROM kardex WHERE cantidad_real>0 and  id_producto='" & rstA("id_producto") & "' and id_alm='00001' and ruc='" & KEY_RUC & "' and id_kardex<'" & rst("id_kardex") & "' and fecha_emision<='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' LIMIT 1"
                 Call ConfiguraRstL(strCadena)
                 If rstL.RecordCount > 0 Then
                    If rst("costo_unitario") <> rstL("costo_promedio") Then
                       
                       
                       
                       strCadena = "UPDATE kardex SET costo_unitario='" & rstL("costo_promedio") & "',costo_promedio='" & rstL("costo_promedio") & "' WHERE id_doc IN('0089','0090','0009','00031') and   id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                       CnBd.Execute (strCadena)
                       
                       If rst("id_doc") = "0009" Or rst("id_doc") = "0031" Then
                          strCadena = "UPDATE movimiento_transferencia_detalle SET precio_costo='" & rstL("costo_promedio") & "' WHERE id_producto='" & rst("id_producto") & "' and  id_transferencia='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                          CnBd.Execute (strCadena)
                       End If
                       
                       If rst("id_doc") = "0089" Or rst("id_doc") = "0090" Then
                          strCadena = "UPDATE movimiento_compra_detalle SET c_unitario='" & rstL("costo_promedio") & "',valor_venta='" & rstL("costo_promedio") * rstL("cantidad") & "',total='" & rstL("costo_promedio") * rstL("cantidad") * 1.18 & "' WHERE id_alm='" & rst("id_alm") & "' and  id_producto='" & rst("id_producto") & "' and  id_compra='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                          CnBd.Execute (strCadena)
                          strCadena = "UPDATE almacen_producto SET precio_compra='" & rstL("costo_promedio") & "' WHERE id_alm='" & rst("id_alm") & "' and  id_producto='" & rst("id_producto") & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                          CnBd.Execute (strCadena)
                       End If
                       
                    End If
                 End If
                'strCadena = "UPDATE kardex SET costo_unitario='" & rst("costo_unitario") / (1.18) & "' WHERE id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                'CnBd.Execute (strCadena)
                '
                rst.MoveNext
            Next j
        Else
            strCadena = "SELECT precio_compra FROM almacen_producto WHERE id_producto='" & rstA("id_producto") & "' and id_alm='00001' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                strCadena = "UPDATE almacen_producto SET precio_compra='" & rstK("precio_compra") & "' WHERE id_producto='" & rstA("id_producto") & "' and id_alm='" & rstA("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
            End If
        End If
       
       rstA.MoveNext
   Next i
End If
Exit Sub

strCadena = "SELECT * FROM kardex WHERE id_alm<>'00001' id_doc='0009' and ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
   X = rst("id_producto")
        strCadena = "UPDATE kardex SET costo_unitario='" & rst("costo_unitario") / (1.18) & "' WHERE id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        strCadena = "UPDATE movimiento_transferencia_detalle SET precio_costo='" & rst("costo_unitario") / (1.18) & "' WHERE id_alm='" & rst("id_alm") & "' and  id_producto='" & rst("id_producto") & "' and  id_transferencia='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If

Exit Sub



strCadena = "SELECT d.precio_costo,d.id_producto,d.id_detalle FROM movimiento_transferencia_detalle d,movimiento_transferencia t  WHERE d.id_transferencia=t.id_transferencia and d.ruc=t.ruc and t.ruc='" & KEY_RUC & "' and t.fecha>='2018-12-01' ORDER BY id_producto ASC  "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prog_indicador.Min = 0
   Me.prog_indicador.Max = rst.RecordCount + 1
   For i = 0 To rst.RecordCount - 1
        If rst("precio_costo") > 0 Then
        in_valor_venta = (rst("precio_costo") / (1 + KEY_IGV))
        strCadena = "UPDATE movimiento_transferencia_detalle SET precio_costo='" & in_valor_venta & "' WHERE id_detalle='" & rst("id_detalle") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        
        strCadena = "UPDATE kardex SET costo_unitario='" & in_valor_venta & "' WHERE id_doc='0009' and  id_producto='" & rst("id_producto") & "' and id_movimiento='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        End If
        
       ' If rst("i") = 1 Then
        '    x = 0
        'End If
        
        
        
        rst.MoveNext
        DoEvents
        Me.prog_indicador.Value = i
        
   Next i
End If


strCadena = "SELECT * FROM view_compra_detalle WHERE    id_doc IN ('0001') and ruc='" & KEY_RUC & "'  ORDER BY id_compra DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prog_indicador.Min = 0
   Me.prog_indicador.Max = rst.RecordCount + 1
   For i = 0 To rst.RecordCount - 1
        If rst("total") > 0 Then
        in_valor_venta = (rst("total") / (1 + KEY_IGV))
        strCadena = "UPDATE movimiento_compra_detalle SET valor_venta='" & in_valor_venta & "' WHERE id_detalle_compra='" & rst("id_detalle_compra") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        in_costo_unitario = in_valor_venta / rst("cantidad")
        
        strCadena = "UPDATE kardex SET costo_unitario='" & in_costo_unitario & "' WHERE cantidad_real>0 and id_alm='" & rst("id_alm") & "' and  id_producto='" & rst("id_producto") & "' and id_movimiento='" & rst("id_compra") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        End If
        
       ' If rst("i") = 1 Then
        '    x = 0
        'End If
        
        
        
        rst.MoveNext
        DoEvents
        Me.prog_indicador.Value = i
        
   Next i
End If




End Sub

Private Sub cmdCuadrarIngresos_Click()
Call put_update_costo_ingreso
Call put_update_costos
Exit Sub


       strCadena = "SELECT * FROM movimiento_compra WHERE   fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0008','0007') and ruc='" & KEY_RUC & "' ORDER BY id_compra ASC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For i = 0 To rst.RecordCount - 1
                
                
                strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_compra") & "' and cantidad_real>0 and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    in_costo_kardex = rstK("costo_kardex")
                Else
                    in_costo_kardex = 0
                    MsgBox "NO ESTA EN EL KARDEX ADMIN" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                in_detalle = rst("serie") & rst("numero")
                strCadena = "SELECT * FROM con_asiento WHERE  IdTipoAsiento='1CIX000000000137' and  glosa LIKE '%" & in_detalle & "%' and IdEmpresaSis='" & KEY_RUC & "' and Activo='1' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    strCadena = "select sum(DebeMN+HaberMN) from con_asientomovimiento m where `m`.`IdCuentaContable` IN('2501000000002212','2501000000002213') and  m.`IdAsiento`='" & rstL("Id") & "' and  Activo='1' and IdEmpresaSis='" & KEY_RUC & "' "
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) <> in_costo_kardex Then
                    
                       MsgBox "DIFERENCIA" + Chr(13) + "ADMIN :" & str(in_costo_kardex) + Chr(13) + "CONTA:" & str(rstL(0))
                       X = 0
                    End If
                Else
                    MsgBox "NO ESTA EN EL CONTABLE" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                
sigu:
Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
DoEvents
                rst.MoveNext
           Next i
        End If


End Sub
Private Sub put_salidas_unico(ByVal in_cuenta_debe As String, ByVal in_cuenta_haber As String)
MsgBox "VERIFICANDO BOLETAS Y FACTURAS", vbInformation

strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0008') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1

      
   
        strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE saldo_stock>=0 and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If IsNull(rstK("costo_kardex")) = True Then
           MsgBox rst("documento") + Chr(13) + "NO SE ENCUENTRA EN EL KARDEX ADMINISTRATIVO", vbInformation
           GoTo sig2
        End If
        
        If rstK.RecordCount > 0 Then
           in_costo_kardex = rstK("costo_kardex")
        Else
            in_costo_kardex = rstK("costo_kardex")
        End If
        in_costo_conta = 0
        in_glosa = Trim(rst("serie") & rst("numero"))
        in_glosa2 = Trim(rst("serie") & "-" & rst("numero"))
        strCadena = "select     (`m`.`DebeMN` + `m`.`HaberMN`) AS `costo`  From  `con_asientomovimiento` `m`   Where  " & _
        " `m`.`Activo` = 1 and `m`.`IdCuentaContable` IN('" & in_cuenta_debe & "','" & in_cuenta_haber & "')  and (m.Glosa LIKE '%" & in_glosa & "%' or m.Glosa LIKE '%" & in_glosa2 & "%' ) LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo_conta = rstL("costo")
        Else
            If IsNull(rstK("costo_kardex")) = True And rstL.RecordCount = 0 Then
                GoTo sig2
            End If
         
             PlaySound App.Path & "\sonidos\dingding.wav"
            ' MsgBox "NO ESTA CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rstK("costo_kardex"), "#,##0.00")
             
             MsgBox "NO ESTA CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rstK("costo_kardex"), "#,##0.00")
             If MsgBox("Desea generar Nuevamente el asiento", vbQuestion + vbYesNo) = vbYes Then
                    
                 If rst("id_tipo") = "02" Then
                    strCadena = "SELECT * FROM view_venta_servicio where servicio='no' and  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstChat(strCadena)
                    If rstChat.RecordCount > 0 Then
                        strCadena = "UPDATE movimiento_venta SET id_tipo='01' WHERE id_venta='" & rst("id_venta") & "'"
                        CnBd.Execute (strCadena)
                    End If
                 End If
             
             
                 strCadena = "call P_insert_venta_agenda_test('" & rst("id_venta") & "')"
                 CnBd.Execute (strCadena)
             End If
             
             GoTo sig2
        End If
     
        If Val(in_costo_conta) <> Val(in_costo_kardex) Then
            MsgBox "INCONSISTENCIA" + Chr(13) + rst("documento") & Space(2) + Chr(13) + Chr(13) + "ADMIN KARDEX:" & Format(rstK("costo_kardex"), "#,##0.00") + Chr(13) + "CONTA KARDEX:" & Format(in_costo_conta, "#,##0.00")
        
            Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & Space(5) & rst("documento")
            If rst("id_doc") = "0007" Then
                strCadena = "call update_costo_venta_vitekey_nota('" & rst("id_venta") & "','" & KEY_RUC & "')"
            Else
                strCadena = "call update_costo_venta_vitekey('" & rst("id_venta") & "','" & KEY_RUC & "')"
            End If
            CnBd.Execute (strCadena)
        End If
        
sig2:
        Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
        rst.MoveNext
        
        DoEvents
   Next i
End If


MsgBox "Verificacion Realziada", vbInformation
Exit Sub

If MsgBox("DESEA UNA VERIFICACION MAS EXAUSTIVA", vbQuestion + vbYesNo) = vbYes Then
        strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0008','0007') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For i = 0 To rst.RecordCount - 1
                strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    in_costo_kardex = rstK("costo_kardex")
                Else
                    in_costo_kardex = 0
                    MsgBox "NO ESTA EN EL KARDEX ADMIN" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                in_detalle = rst("serie") & rst("numero")
                strCadena = "SELECT * FROM con_asiento WHERE  IdTipoAsiento='1CIX000000000137' and  glosa LIKE '%" & in_detalle & "%' and IdEmpresaSis='" & KEY_RUC & "' and Activo='1' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    strCadena = "select sum(DebeMN+HaberMN) from con_asientomovimiento m where `m`.`IdCuentaContable` IN('2501000000002212','2501000000002213') and  m.`IdAsiento`='" & rstL("Id") & "' and  Activo='1' and IdEmpresaSis='" & KEY_RUC & "' "
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) <> in_costo_kardex Then
                    
                       MsgBox "DIFERENCIA" + Chr(13) + "ADMIN :" & str(in_costo_kardex) + Chr(13) + "CONTA:" & str(rstL(0))
                       X = 0
                    End If
                Else
                    MsgBox "NO ESTA EN EL CONTABLE" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                
sigu:
Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
DoEvents
                rst.MoveNext
           Next i
        End If
        
End If



End Sub
Private Function get_id_cuenta(ByVal in_cuenta As String) As String
strCadena = "SELECT Id FROM con_cuentacontable WHERE Ejercicio='" & Year(Me.DtpDesde.Value) & "' and  NroCuenta='" & in_cuenta & "' and IdEmpresaSis='" & KEY_RUC & "'"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   get_id_cuenta = rstA("id")
End If
End Function
Private Sub cmdCuadrarSalidas_Click()

Dim in_cuenta_debe As String
Dim in_cuenta_haber As String



If KEY_RUC = "20128836251" Then
   in_cuenta_debe = get_id_cuenta("6910101")
   in_cuenta_haber = get_id_cuenta("6910102")
Else
    
    
    If KEY_RUC = "20493910052" Then
        in_cuenta_debe = "0501000000003998"
        in_cuenta_haber = "0501000000002692"
    End If
    
    If KEY_RUC = "20487725286" Then
        in_cuenta_debe = "2801000000006539"
        in_cuenta_haber = "2801000000005181"
    End If
    
    
    
    Call put_salidas_unico(in_cuenta_debe, in_cuenta_haber)
    Exit Sub
End If














'***** VERIFICACION DE ANULADOS
MsgBox "VERIFICANDO SI EXISTEN COMPROBANTES ANULADOS", vbInformation
strCadena = "SELECT * FROM movimiento_venta WHERE anulado='si' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0007','0008') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If IsNull(rstK("costo_kardex")) = False Then
               in_costo_kardex = rstK("costo_kardex")
           
           
        Else
            in_costo_kardex = 0
        End If
        in_costo_conta = 0
        
        strCadena = "select     (`m`.`DebeMN` + `m`.`HaberMN`) AS `costo`  From     ((`con_documento` `d` join `con_asiento` `a`) join `con_asientomovimiento` `m`) " & _
        "   Where     ((`d`.`Id` = `a`.`IdReferencia`) and (`a`.`Id` = `m`.`IdAsiento`) and (`d`.`IdEmpresaSis` = `a`.`IdEmpresaSis`) and (`a`.`IdEmpresaSis` = `m`.`IdEmpresaSis`) and (`d`.`Activo` = 1) and (`a`.`Activo` = 1) and " & _
        "(`m`.`Activo` = 1) and (`m`.`IdCuentaContable` IN('" & in_cuenta_debe & "','" & in_cuenta_haber & "')) AND d.IdReferencia='" & rst("id_venta") & "') and d.IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        
        
        If IsNull(rstL("costo")) = False Then
            
            GoTo siguiente
        Else
            If IsNull(rstK("costo_kardex")) = True And rstL.RecordCount = 0 Then
                GoTo siguiente
            End If
             strCadena = "UPDATE kardex SET ruc='0' WHERE id_movimiento='" & rst("id_venta") & "' and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and id_doc='" & rst("id_doc") & "' and id_numero='" & rst("numero") & "' and ruc='" & KEY_RUC & "'"
             CnBd.Execute (strCadena)
             PlaySound App.Path & "\sonidos\dingding.wav"
             MsgBox "NO ESTA EN CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rst("total"), "#,##0.00")
            GoTo siguiente
        End If
        
        
        
        If Val(in_costo_conta) > 0 Then
            MsgBox "ANULADO EN ADMINISTRATIVO" + Chr(13) + rst("documento") & Space(2) & Format(rst("total"), "#,##0.00")
        End If
siguiente:
         Me.Command1.Caption = i & Space(2) & rst("fecha_emision")
        rst.MoveNext
       
        DoEvents
   Next i
End If

MsgBox "VERIFICANDO NOTAS DE CREDITO", vbInformation
strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0007') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT IFNULL(round(sum(cantidad*costo_unitario),4),0) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           in_costo_kardex = rstK("costo_kardex")
        Else
            in_costo_kardex = rstK("costo_kardex")
        End If
        in_costo_conta = 0
        strCadena = "select     (`m`.`DebeMN` + `m`.`HaberMN`) AS `costo`  From     ((`con_documento` `d` join `con_asiento` `a`) join `con_asientomovimiento` `m`) " & _
        "   Where     ((`d`.`Id` = `a`.`IdReferencia`) and (`a`.`Id` = `m`.`IdAsiento`) and (`d`.`IdEmpresaSis` = `a`.`IdEmpresaSis`) and (`a`.`IdEmpresaSis` = `m`.`IdEmpresaSis`) and (`d`.`Activo` = 1) and (`a`.`Activo` = 1) and " & _
        "(`m`.`Activo` = 1) and (`m`.`IdCuentaContable` IN('" & in_cuenta_debe & "','" & in_cuenta_haber & "')) AND d.IdReferencia='" & rst("id_venta") & "') and d.IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo_conta = rstL("costo")
        Else
            If IsNull(rstK("costo_kardex")) = True And rstL.RecordCount = 0 Then
                GoTo sig3
            End If
         
             PlaySound App.Path & "\sonidos\dingding.wav"
             MsgBox "NO ESTA CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rstK("costo_kardex"), "#,##0.00")
             If MsgBox("Desea generar Nuevamente el asiento", vbQuestion + vbYesNo) = vbYes Then
                 strCadena = "call P_insert_venta_asiento_contable('" & rst("id_venta") & "')"
                 CnBd.Execute (strCadena)
             End If
             
             
             GoTo sig3
        End If
     
        If Val(in_costo_conta) <> Val(in_costo_kardex) Then
            MsgBox "INCONSISTENCIA" + Chr(13) + rst("documento") & Space(2) + "ADMIN KARDEX:" & Format(rstK("costo_kardex"), "#,##0.0000") + Chr(13) + "CONTA KARDEX:" & Format(in_costo_conta, "#,##0.0000")
        
            Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & Space(5) & rst("documento")
            If rst("id_doc") = "0007" Then
                strCadena = "call update_costo_venta_vitekey_nota('" & rst("id_venta") & "','" & KEY_RUC & "')"
            Else
                strCadena = "call update_costo_venta_vitekey('" & rst("id_venta") & "','" & KEY_RUC & "')"
            End If
            CnBd.Execute (strCadena)
        End If
        
sig3:
        Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
        rst.MoveNext
        
        DoEvents
   Next i
End If


        
        
        
        
        
        
        
        







MsgBox "VERIFICANDO BOLETAS Y FACTURAS", vbInformation

strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0008') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
s:
       If i < 477 Then
        rst.MoveNext
         i = i + 1
         GoTo s
       End If
   
        strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If IsNull(rstK("costo_kardex")) = True Then
           MsgBox rst("documento") + Chr(13) + "NO SE ENCUENTRA EN EL KARDEX ADMINISTRATIVO", vbInformation
           GoTo sig2
        End If
        
        If rstK.RecordCount > 0 Then
           in_costo_kardex = rstK("costo_kardex")
        Else
            in_costo_kardex = rstK("costo_kardex")
        End If
        in_costo_conta = 0
        strCadena = "select     (`m`.`DebeMN` + `m`.`HaberMN`) AS `costo`  From     ((`con_documento` `d` join `con_asiento` `a`) join `con_asientomovimiento` `m`) " & _
        "   Where     ((`d`.`Id` = `a`.`IdReferencia`) and (`a`.`Id` = `m`.`IdAsiento`) and (`d`.`IdEmpresaSis` = `a`.`IdEmpresaSis`) and (`a`.`IdEmpresaSis` = `m`.`IdEmpresaSis`) and (`d`.`Activo` = 1) and (`a`.`Activo` = 1) and " & _
        "(`m`.`Activo` = 1) and (`m`.`IdCuentaContable` IN('" & in_cuenta_debe & "','" & in_cuenta_haber & "')) AND d.IdReferencia='" & rst("id_venta") & "') and d.IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo_conta = rstL("costo")
        Else
            If IsNull(rstK("costo_kardex")) = True And rstL.RecordCount = 0 Then
                GoTo sig2
            End If
         
             PlaySound App.Path & "\sonidos\dingding.wav"
            ' MsgBox "NO ESTA CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rstK("costo_kardex"), "#,##0.00")
             
             MsgBox "NO ESTA CONTABILIDAD" + Chr(13) + rst("documento") & Space(2) & Format(rstK("costo_kardex"), "#,##0.00")
             If MsgBox("Desea generar Nuevamente el asiento", vbQuestion + vbYesNo) = vbYes Then
                    
                 If rst("id_tipo") = "02" Then
                    strCadena = "SELECT * FROM view_venta_servicio where servicio='no' and  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstChat(strCadena)
                    If rstChat.RecordCount > 0 Then
                        strCadena = "UPDATE movimiento_venta SET id_tipo='01' WHERE id_venta='" & rst("id_venta") & "'"
                        CnBd.Execute (strCadena)
                    End If
                 End If
             
             
                 strCadena = "call P_insert_venta_asiento_contable('" & rst("id_venta") & "')"
                 CnBd.Execute (strCadena)
             End If
             
             GoTo sig2
        End If
        If IsNull(in_costo_conta) = True Then
            in_costo_conta = 0
        End If
        
        If Val(in_costo_conta) <> Val(in_costo_kardex) Then
            MsgBox "INCONSISTENCIA" + Chr(13) + rst("documento") & Space(2) + Chr(13) + Chr(13) + "ADMIN KARDEX:" & Format(rstK("costo_kardex"), "#,##0.00") + Chr(13) + "CONTA KARDEX:" & Format(in_costo_conta, "#,##0.00")
        
            Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & Space(5) & rst("documento")
            If rst("id_doc") = "0007" Then
                strCadena = "call update_costo_venta_vitekey_nota('" & rst("id_venta") & "','" & KEY_RUC & "')"
            Else
                strCadena = "call update_costo_venta_vitekey('" & rst("id_venta") & "','" & KEY_RUC & "')"
            End If
            CnBd.Execute (strCadena)
        End If
        
sig2:
        Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
        rst.MoveNext
        
        DoEvents
   Next i
End If


MsgBox "Verificacion Realziada", vbInformation


If MsgBox("DESEA UNA VERIFICACION MAS EXAUSTIVA", vbQuestion + vbYesNo) = vbYes Then
        strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0001','0003','0008','0007') and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For i = 0 To rst.RecordCount - 1
                strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    in_costo_kardex = rstK("costo_kardex")
                Else
                    in_costo_kardex = 0
                    MsgBox "NO ESTA EN EL KARDEX ADMIN" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                in_detalle = rst("serie") & rst("numero")
                strCadena = "SELECT * FROM con_asiento WHERE  IdTipoAsiento='1CIX000000000137' and  glosa LIKE '%" & in_detalle & "%' and IdEmpresaSis='" & KEY_RUC & "' and Activo='1' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    strCadena = "select sum(DebeMN+HaberMN) from con_asientomovimiento m where `m`.`IdCuentaContable` IN('2501000000002212','2501000000002213') and  m.`IdAsiento`='" & rstL("Id") & "' and  Activo='1' and IdEmpresaSis='" & KEY_RUC & "' "
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) <> in_costo_kardex Then
                    
                       MsgBox "DIFERENCIA" + Chr(13) + "ADMIN :" & str(in_costo_kardex) + Chr(13) + "CONTA:" & str(rstL(0))
                       X = 0
                    End If
                Else
                    MsgBox "NO ESTA EN EL CONTABLE" + Chr(13) + rst("documento")
                    GoTo sigu
                End If
                
                
sigu:
Me.cmdCuadrarSalidas.Caption = str(i) & Space(2) & str(rst.RecordCount - 1) & Space(5) & rst("fecha_emision")
DoEvents
                rst.MoveNext
           Next i
        End If
        
End If




End Sub

Private Sub CmdEjecutar_Click()


End Sub


Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
'  Grilla.Clear
  
 '  Grilla.Rows = Rst.RecordCount
   
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 1000
  Grilla.ColWidth(2) = 2400
  Grilla.ColWidth(3) = 2000
  Grilla.ColWidth(4) = 2000
  Grilla.ColWidth(5) = 2200
  Grilla.ColWidth(6) = 1200
  Grilla.ColWidth(7) = 0
  Grilla.ColWidth(8) = 0
  Grilla.ColWidth(9) = 0
  Grilla.ColWidth(10) = 0
  
Call DarFormatoFecha(Grilla, 1)
Call DarFormato(Grilla, 3)
Call DarFormato(Grilla, 4)
Call DarFormato(Grilla, 5)
Call DarFormato(Grilla, 6)
Call DarFormato(Grilla, 7)
Call DarFormato(Grilla, 8)

Set rst = Nothing

  
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub imorimir_matricial()


End Sub



Private Sub cmdInventario_valorizadoGeneral_Click()

End Sub

Private Sub cmdInventarioValorizado_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = KEY_EMPRESA
arr(1, 2) = KEY_DIRECCION

param = arr()
Call impresion_kardex_valorizado_demo(Me.DtpDesde.Value, Me.DtpHasta.Value, Me.DtcAlmacen.BoundText, Trim(Me.TxtcodigoProd.Text))



'Call impresion_kardex_valorizado(Me.DtpDesde.Value, Me.DtpHasta.Value, Me.DtcAlmacen.BoundText)
Exit Sub

 strCadena = "SELECT 'PERIODO',k.`ruc`,k.`id_producto`,'01-MERCADERIA',p.`nombre_prod`,k.`fecha_emision`,LPAD(CAST(k.id_doc AS integer),2,0),k.`id_serie`,k.`id_numero`," & _
 " k.`id_tipo_movimiento`,k.`cantidad`,k.`cantidad_real`,k.`costo_unitario`,k.`saldo_stock` From tmp_kardex k, producto p Where " & _
 " k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.ruc='" & KEY_RUC & "' and   k.`id_producto`=p.`id_producto` and k.`ruc`=p.`ruc` ORDER BY k.`fecha_emision` ASC,k.`id_kardex` ASC "
 Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "rpt_kardex_valorizado", param, App.Path + "\Reportes\")
 
Exit Sub
 strCadena = "call put_crear_kardex_temporal('" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "');"
 CnBd.Execute (strCadena)
 
 strCadena = "call cursor_stock_inicial('" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_RUC & "');"
 CnBd.Execute (strCadena)
 
 strCadena = "SELECT 'PERIODO',k.`ruc`,k.`id_producto`,'01-MERCADERIA',p.`nombre_prod`,k.`fecha_emision`,k.`id_doc`,k.`id_serie`,k.`id_numero`," & _
 " k.`id_tipo_movimiento`,k.`cantidad`,k.`cantidad_real`,k.`costo_unitario`,k.`saldo_stock` From tmp_kardex k, producto p Where " & _
 " k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.ruc='" & KEY_RUC & "' and   k.`id_producto`=p.`id_producto` and k.`ruc`=p.`ruc` ORDER BY k.`fecha_emision` ASC,k.`id_kardex` ASC "
 Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "rpt_kardex_valorizado", param, App.Path + "\Reportes\")
 
 Exit Sub
 
 
 
 
 strCadena = "SELECT * FROM tmp_kardex WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_tipo_movimiento='06' and id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "' order by id_producto ASC "
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    rst.MoveFirst
   
    For i = 0 To rst.RecordCount - 1
        in_costo_promedio = 0
        strCadena = "SELECT IFNULL(costo_promedio,0)  FROM tmp_kardex WHERE fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
        
        Set Me.HfdGrilla.Recordset = rstK
        in_costo_promedio = rstK(0)
        
       
        strCadena = "UPDATE tmp_kardex SET costo_unitario='" & in_costo_promedio & "',costo_promedio='" & in_costo_promedio & "' WHERE id_kardex='" & rst("id_kardex") & "'"
        CnBd.Execute (strCadena)
        strCadena = "DELETE FROM tmp_kardex WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' and fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "'"
        CnBd.Execute (strCadena)
        DoEvents
       ' Me.cmdInventario_valorizadoGeneral.Caption = str(i) & Space(2) & rst.RecordCount
        
        rst.MoveNext
    Next i
    
 End If
sss:
 strCadena = "SELECT 'PERIODO',k.`ruc`,k.`id_producto`,'01-MERCADERIA',p.`nombre_prod`,k.`fecha_emision`,k.`id_doc`,k.`id_serie`,k.`id_numero`," & _
 " k.`id_tipo_movimiento`,k.`cantidad`,k.`cantidad_real`,k.`costo_unitario`,k.`saldo_stock` From tmp_kardex k, producto p Where " & _
 " k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.ruc='" & KEY_RUC & "' and   k.`id_producto`=p.`id_producto` and k.`ruc`=p.`ruc` ORDER BY k.`fecha_emision` ASC,k.`id_kardex` ASC "
 Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "rpt_kardex_valorizado", param, App.Path + "\Reportes\")
 Exit Sub
 
 strCadena = "SELECT DISTINCT id_producto FROM tmp_kardex WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Set Me.HfdGrilla.Recordset = rst
 End If
 Exit Sub
 
 
 strCadena = "SELECT * FROM tmp_kardex WHERE id_tipo_movimiento='06' and fecha_emision='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Set Me.HfdGrilla.Recordset = rst
 End If
 
 
 
 
 
 
Exit Sub






 '*****  TEMPORAL KARDEX

 
 
 
 Exit Sub
 
 
 

        
   

If Trim(Me.TxtcodigoProd.Text) = "" Then
    strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE  k.id_producto=p.id_producto and k.ruc=p.ruc and k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' ORDER BY k.id_producto"
Else
    strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' and k.id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' ORDER BY k.id_producto"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
        strCadena = "call procedure_kardex_inicial('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRstK(strCadena)
        
       ' strCadena = "CALL cursor_kardex_valorizado('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & rstK("cantidad_inicial") & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       ' CnBd.Execute (strCadena)
        
        
        rst.MoveNext
        Me.prg_avance.Value = i
        Me.cmdKardexGeneral.Caption = (i + 1) & Space(2) & "-" & Space(1) & rst.RecordCount
        DoEvents
   Next i
End If
   
   
strCadena = "SELECT * FROM tmp_kardex where id_producto='00007'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    MsgBox rst()
End If
   
End Sub

Private Sub cmdKardexContable_Click()


Procedencia = nuevo
frmTempExportExcel.Show
Exit Sub



Dim cam(0 To 1, 1 To 2)  As String
cam(0, 1) = "productodes"
cam(1, 1) = "almacendes"

cam(0, 2) = Trim(Me.DtcProducto.Text)
cam(1, 2) = Trim(Me.DtcAlmacen.Text)

param = cam()


                   
strCadena = "SELECT id_kardex,fecha_emision,comprobante,id_persona,ncliente,cantidad,cantidad_real,round(costo_unitario,3),saldo_stock,round(costo_promedio,3) FROM view_kardex WHERE id_tipo_movimiento<>'10' and  id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptKardex_producto", param, App.Path + "\Reportes\")
   Exit Sub
   
   

End Sub



Private Sub cmdKardexGeneral_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY") & Space(3) & Format(Me.DtpHasta.Value, "dd-mm-YYYY ")
arr(1, 2) = get_saldo_inicial

param = arr()

If Me.chk_sinprocesar.Value = 1 Then
     strCadena = "SELECT id_producto,producto,cantidad_inicial,saldo_inicial,cantidad_ingreso,saldo_ingreso,cantidad_salida,saldo_salida,cantidad_final,round(saldo_final,2) FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)

    Ans = ShowMultiReport(rst, "RptKardexValorizado", param, App.Path + "\Reportes\")
    Exit Sub
End If

strCadena = "DELETE FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



If Trim(Me.TxtcodigoProd.Text) = "" Then
    strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE p.id_linea not in('00009','00017') and  k.id_producto=p.id_producto and k.ruc=p.ruc and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' ORDER BY k.id_producto"
Else
    strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   k.id_producto not in('00000','00') and k.ruc='" & KEY_RUC & "' and k.id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' ORDER BY k.id_producto"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rst.RecordCount
   
   strCadena = "call put_crear_kardex_temporal('" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "');"
   CnBd.Execute (strCadena)
   
   
   For i = 0 To rst.RecordCount - 1
        strCadena = "CALL procedure_kardex_general('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        Me.prg_avance.Value = i
        Me.cmdKardexGeneral.Caption = (i + 1) & Space(2) & "-" & Space(1) & rst.RecordCount
        DoEvents
   Next i

End If



strCadena = "SELECT id_producto,producto,cantidad_inicial,saldo_inicial,cantidad_ingreso,saldo_ingreso,cantidad_salida,saldo_salida,cantidad_final,saldo_final FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptKardexValorizado", param, App.Path + "\Reportes\")


End Sub
Private Function get_saldo_inicial() As Double
'saldo incial
If KEY_RUC = "20128836251" Then
    strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k WHERE k.`fecha_emision`<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_cantidad_inicial = rst(0)
    in_saldo_inicial = rst(1)
Else
    If Month(Me.DtpDesde.Value) = 1 Then
        strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k WHERE k.`fecha_emision`<='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        in_cantidad_inicial = rst(0)
        in_saldo_inicial = rst(1)
    Else
        strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k WHERE k.`fecha_emision`<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        in_cantidad_inicial = rst(0)
        in_saldo_inicial = rst(1)
    End If
End If
get_saldo_inicial = in_saldo_inicial
End Function
Private Sub cmdKardexResumido_Click()
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")

param = arr()


'saldo incial
If KEY_RUC = "20128836251" Then
    
    strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k WHERE k.`fecha_emision`<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_cantidad_inicial = rst(0)
    in_saldo_inicial = rst(1)
  
Else
    If Month(Me.DtpDesde.Value) = 1 Then
        strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k,producto p WHERE  k.id_producto=p.id_producto and k.ruc=p.ruc and  p.id_linea not in('00009','00017') and   k.`fecha_emision`<='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        in_cantidad_inicial = rst(0)
        in_saldo_inicial = rst(1)
    Else
        strCadena = "SELECT SUM(k.`cantidad_real`),sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  p.id_linea not in('00009','00017') and k.`fecha_emision`<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        in_cantidad_inicial = rst(0)
        in_saldo_inicial = rst(1)
    End If
End If



'ingresos
If KEY_RUC = "20128836251" Then
    strCadena = "SELECT sum(k.`cantidad_real`*k.`costo_unitario`) as ingresos FROM kardex k WHERE k.`fecha_emision`>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`fecha_emision`<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and k.`cantidad_real`>0 and k.`ruc`='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_saldo_ingreso = rst(0)
Else
    strCadena = "SELECT sum(k.`cantidad_real`*k.`costo_unitario`) as ingresos FROM kardex k WHERE k.id_doc not in('0089','0009','0054','0097','0031','0106')and k.`saldo_stock` >= 0 and k.`costo_promedio` >= 0 and  k.`fecha_emision`>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`fecha_emision`<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and k.`cantidad_real`>0 and k.`ruc`='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_saldo_ingreso = rst(0)
End If
'salidas
If KEY_RUC = "20128836251" Then
    strCadena = "SELECT sum(k.`cantidad_real`*k.`costo_promedio`) as salidas FROM kardex k ,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  p.id_linea not in('00009','00017') and    k.`fecha_emision`>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`fecha_emision`<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and k.`cantidad_real`<0 and k.`ruc`='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    in_saldo_salida = rst(0)
    in_salidas = Abs(rst(0))
Else
    strCadena = "SELECT sum(k.`cantidad_real`*k.`costo_promedio`) as salidas FROM kardex k,producto p WHERE k.id_producto=p.id_producto and k.ruc=p.ruc and  p.id_linea not in('00009','00017') and  k.`id_doc` not in ('0089','0009','0054','0097','0031') and k.`saldo_stock` >= 0 and k.`costo_promedio` >= 0 and  k.`fecha_emision`>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`fecha_emision`<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and k.`cantidad_real`<0 and k.`ruc`='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    in_saldo_salida = rst(0)
    in_salidas = Abs(rst(0))
End If
in_saldo_final = in_saldo_inicial + in_saldo_ingreso + in_saldo_salida

strCadena = "DELETE FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "INSERT INTO kardex_valorizado_sunat(`id_producto`,`producto`,cantidad_inicial,`saldo_inicial`,`saldo_ingreso`,`saldo_salida`,`saldo_final`,`dni_save`,`ruc`)VALUES " & _
"('00','RESUMEN VALORIZADO','" & in_cantidad_inicial & "','" & in_saldo_inicial & "','" & in_saldo_ingreso & "','" & in_salidas & "','" & in_saldo_final & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT id_producto,producto,cantidad_inicial,saldo_inicial,saldo_ingreso,saldo_salida,saldo_final FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)


Ans = ShowMultiReport(rst, "RptKardexResumen", param, App.Path + "\Reportes\")

End Sub

Private Sub cmdreportekardex_Click()
Dim cam(0 To 1, 1 To 2)  As String



    cam(0, 1) = "productodes"
    cam(1, 1) = "almacendes"

    cam(0, 2) = Trim(Me.DtcProducto.Text)
    cam(1, 2) = Trim(Me.DtcAlmacen.Text)

param = cam()


                   
strCadena = "SELECT id_kardex,fecha_emision,comprobante,id_persona,ncliente,cantidad,cantidad_real,round(costo_unitario,3),saldo_stock,round(costo_promedio,3) FROM view_kardex WHERE id_tipo_movimiento<>'10' and  id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptKardex_producto", param, App.Path + "\Reportes\")
   Exit Sub
End Sub

Private Sub cmdreportevalorizado_Click()
If Exportar_Excel(App.Path & "\excel\kardex.xls", Me.HfdGrilla) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub
Private Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  On Error GoTo error_Handler
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows
            For columna = 1 To .Cols - 1
                
                    o_Hoja.Cells(Fila, columna + 1).Value = .TextMatrix(Fila - 1, columna)
              
            Next
        Next
    End With
    o_Libro.close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function
' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub















Private Sub cmdsaldoInicial_Click()
Dim in_producto As String
strCadena = "SELECT * FROM `inventario_vargas_31122017"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
   
   in_producto = Format(rst("id_producto"), "00000")
   strCadena = "SELECT id_kardex FROM kardex WHERE id_doc='0106' and cantidad='" & rst("cantidad") & "' and  fecha_emision='2017-12-31'  and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount < 1 Then
        MsgBox "PRODUCTO :" & in_producto, vbInformation
        Call update_saldo_inicial_vargas(in_producto)
 
   End If
   rst.MoveNext
   Me.Command1.Caption = i
   DoEvents
   Next i
End If
MsgBox "ya"

Exit Sub
End Sub

Private Sub cmdupdate_Click()

If KEY_RUC = "20128836251" Then
    Call update_kardex_VARGAS(Trim(Me.TxtcodigoProd.Text))
    
Else
    'Call update_kardex_v2(Trim(Me.txtcodigoprod.Text))
    Call put_kardex_temporal(Trim(Me.TxtcodigoProd.Text))
End If




End Sub



Public Sub put_kardex_temporal(ByVal in_producto As String)


strCadena = "SELECT * FROM producto WHERE  id_producto='" & in_producto & "' and   ruc='" & KEY_RUC & "' ORDER BY id_producto  ASC"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   Me.prog_indicador.Min = 0
   Me.prog_indicador.Max = rstIN.RecordCount

    DoEvents
    Me.cmdUpdateKardexTotal.Caption = str(i) & Space(2) & str(rstIN.RecordCount)
    Me.TxtcodigoProd.Text = rstIN("id_producto")
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
    If Me.chk_periodo.Value = 1 Then
       in_fecha_kardex = Format(get_periodo_fecha_ini(Me.DtcPeriodo.BoundText), "YYYY-mm-dd")
   Else
      in_fecha_kardex = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
   End If
    
  
    
    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha_kardex, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
       
    Call put_kardex_tipo(Trim(Me.TxtcodigoProd.Text), in_fecha_kardex)

    DoEvents
    prog_indicador.Value = i

   MsgBox "Proceso Completo", vbInformation
End If
End Sub


Private Sub update_kardex_v2(ByVal in_producto As String)
    Dim in_dias As Integer
    Dim in_flag As Boolean
    Dim in_fecha_kardex As String
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
    If Me.chkBuscarfechas.Value = 1 Then
        in_fecha_kardex = Format(Me.DtpDesde.Value, "YYYY-mm-dd")
    Else
        in_fecha_kardex = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
    End If
    
    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha_kardex, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    
    If Me.chkBuscarfechas.Value = 1 Then
            in_fechai = in_fecha_kardex
    Else
        If Format(get_fecha_periodo_abierto, "YYYY-mm-dd") = "2018-01-01" Then
            Call update_saldo_inicial(in_producto)
            in_fechai = "2018-01-01"
        Else
            in_fechai = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
        End If
    End If
    
    
            
            
            
            
            
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            Me.prog_indicador.Min = 0
            prog_indicador.Max = in_dias + 1
            
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                
             
                
                'COMPRAS
                
            If KEY_PAIS <> KEY_PERU Then
                in_flag = False
                strCadena = "SELECT * FROM movimiento_compra_detalle WHERE   id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount < 1 Then
                   in_flag = True
                
                Else
                   in_flag = False
                   GoTo empezar
                End If

                
                strCadena = "select * from view_transferencia_existe WHERE  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount < 1 Then
                    in_flag = True
                Else
                    in_flag = False
                    GoTo empezar
                End If
                
                strCadena = "select * from movimiento_venta_detalle WHERE id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount < 1 Then
                     in_flag = True
                Else
                     in_flag = False
                     GoTo empezar
                End If
                
                
               
                
                
                If in_flag = True Then
                    GoTo siguiente_flag
                End If
                
            End If
empezar:
                
                
                
                strCadena = "select * from view_kardex_compra_existe WHERE fecha_kardex='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If KEY_RUC = "20128836251" Then
                        Call compras_producto_vargas(in_fechai, in_producto)
                    Else
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                End If
               ' transferencias ingreso
               If KEY_PAIS = KEY_PERU Then
                strCadena = "select id_transferencia from movimiento_transferencia WHERE fecha='" & Format(in_fechai, "YYYY-mm-dd") & "'and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstZ(strCadena)
                If rstZ.RecordCount > 0 Then
                
                strCadena = "select * from view_transferencia_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                  Call transferencia_ingreso_producto(in_fechai, in_producto)
                  Call transferencia_salida_producto(in_fechai, in_producto)
                End If
               End If
               ' DoEvents
               End If
                'ventas salida
               
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                'DoEvents
                'notas salida
                
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                End If
                DoEvents
                prog_indicador.Value = m
                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
siguiente_flag:
       'MsgBox "PROCESADO.....", vbInformation, KEY_VENDEDOR
        


End Sub
Private Sub cmdValorizado_Click()

End Sub

Private Sub cmdUpdateKardexTotal_Click()

strCadena = "SELECT * FROM producto WHERE   ruc='" & KEY_RUC & "' ORDER BY id_producto  ASC"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   
   If Me.chk_periodo.Value = 1 Then
       in_fecha_kardex = Format(get_periodo_fecha_ini(Me.DtcPeriodo.BoundText), "YYYY-mm-dd")
   Else
      in_fecha_kardex = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
   End If
   
   Me.prog_indicador.Min = 0
   Me.prog_indicador.Max = rstIN.RecordCount
   For i = 0 To rstIN.RecordCount - 1
    DoEvents
    Me.cmdUpdateKardexTotal.Caption = str(i) & Space(2) & str(rstIN.RecordCount)
    Me.TxtcodigoProd.Text = rstIN("id_producto")
    DoEvents
    
    
    If verificacion_servicio(Me.TxtcodigoProd.Text) = True Then
        Exit Sub
    End If
    
    
    
  

    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha_kardex, "YYYY-mm-dd") & "' and    id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    
    
    
    
    Call put_kardex_tipo(Trim(Me.TxtcodigoProd.Text), in_fecha_kardex)
    rstIN.MoveNext
    DoEvents
    Me.cmdUpdateKardexTotal.Caption = str(i) & Space(5) & str(rstIN.RecordCount)
    prog_indicador.Value = i
    
   Next i
   MsgBox "Proceso Completo", vbInformation
End If


End Sub

Private Sub carga_inicial(ByVal in_producto As String, ByVal in_alm As String)
strCadena = "SELECT id_kardex, costo_promedio FROM kardex where costo_promedio<1 and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC, id_kardex ASC"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    rstZ.MoveFirst
For j = 0 To rstZ.RecordCount - 1
    If rstZ("costo_promedio") < 1 Then
        strCadena = "SELECT * FROM almacen_producto WHERE  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' ORDER BY precio_compra DESC LIMIT 1"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            strCadena = "UPDATE almacen_producto SET precio_compra='" & rstT("precio_compra") & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
            
            strCadena = "UPDATE kardex SET costo_unitario='" & rstT("precio_compra") & "',costo_promedio='" & rstT("precio_compra") & "' WHERE id_kardex='" & rstZ("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
        End If
        
        
    End If

    rstZ.MoveNext
Next j
End If
End Sub
Private Sub carga_inicial_demo(ByVal in_producto As String, ByVal in_alm As String)

        strCadena = "SELECT * FROM almacen_producto WHERE  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' ORDER BY precio_compra DESC LIMIT 1"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            strCadena = "UPDATE almacen_producto SET precio_compra='" & rstT("precio_compra") & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
            
            
        End If
        
        
 
End Sub


Private Sub put_verificar_kardex_producto(ByVal in_producto As String)

If Val(in_producto) > 0 Then
    strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "' ORDER BY id_alm ASC"
Else
    strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex WHERE  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
End If

Call ConfiguraRstC(strCadena)
If rstc.RecordCount > 0 Then
   rstc.MoveFirst
   Me.prog_indicador.Min = 1
   Me.prog_indicador.Max = rstc.RecordCount
   
   For i = 1 To rstc.RecordCount
        Me.TxtcodigoProd.Text = rstc("id_producto")
        Call put_update_saldo_stock_all(rstc("id_producto"), rstc("id_alm"))
        DoEvents
        Me.prog_indicador.Value = i
        Me.Command1.Caption = str(i) & Space(2) & rstc("id_producto") & Space(2) & rstc("id_alm") & Space(2) & rstc.RecordCount
        rstc.MoveNext
        
        
       
          DoEvents
   Next i
   MsgBox "LISTO"
End If


End Sub

Private Sub put_verificar_kardex(ByVal in_alm As String)

strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
Call ConfiguraRstC(strCadena)
If rstc.RecordCount > 0 Then
   rstc.MoveFirst
   Me.prog_indicador.Min = 1
   Me.prog_indicador.Max = rstc.RecordCount
   
   For i = 1 To rstc.RecordCount
        Me.TxtcodigoProd.Text = rstc("id_producto")
        Call put_update_saldo_stock_all(rstc("id_producto"), rstc("id_alm"))
        rstc.MoveNext
        DoEvents
        
        Me.prog_indicador.Value = i
        Me.Command1.Caption = str(i) & Space(2) & rstc("id_producto") & Space(2) & rstc("id_alm") & Space(2) & rstc.RecordCount
          DoEvents
   Next i
   MsgBox "LISTO"
End If


End Sub


Private Sub cmreportedetallado_Click()

strCadena = "call ADM_kardex_report('1','" & Me.TxtcodigoProd.Text & "','" & Me.DtcAlmacen.BoundText & "','0','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        
         MsgBox str(rst("id_compra")) & Space(2) + rst("comprobante")
         
         rst.MoveNext
    Next i
End If



strCadena = "call ADM_kardex_report('1','" & Me.TxtcodigoProd.Text & "','" & Me.DtcAlmacen.BoundText & "','0','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        
         MsgBox str(rst("id_compra")) & Space(2) + rst("comprobante")
         
         rst.MoveNext
    Next i
End If





End Sub

Private Sub Command1_Click()
Dim in_cantidad As Double
Dim in_costo As Double
Dim in_costo_ant As Double
Dim in_totalk As Double
Dim in_totalt As Double
Dim in_ini As Double
Dim in_cant_ini As Single
Dim in_cant_ing As Single
Dim in_cant_sal As Single
Dim in_saldo_stock As Single
Dim in_ing As Double
Dim in_sal As Double
Dim in_saldo As Double


Call put_verificar_kardex(Me.DtcAlmacen.BoundText)


Exit Sub




strCadena = "SELECT  id_producto,id_alm FROM almacen_producto WHERE  precio_compra<1 and  ruc='" & KEY_RUC & "' ORDER By id_producto ASC"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   For i = 0 To rstA.RecordCount - 1
       Me.DtcAlmacen.BoundText = rstA("id_alm")
       Me.TxtcodigoProd.Text = rstA("id_producto")
        Call carga_inicial_demo(rstA("id_producto"), rstA("id_alm"))
        'Call load_kardex
        'Call put_costo_promedio_almacen(rstA("id_producto"), rstA("id_alm"))
        DoEvents
        rstA.MoveNext
        Me.Command1.Caption = str(i) & Space(2) & str(rstA.RecordCount)
   Next i
End If

MsgBox "listo"

Exit Sub






strCadena = "SELECT * FROM movimiento_venta v, movimiento_venta_detalle d  WHERE v.id_venta=d.id_venta and v.ruc=d.ruc and d.id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and v.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and v.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and v.ruc='" & KEY_RUC & "' and v.id_doc In('0001','0003','0007') and v.ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "SELECT * FROM kardex WHERE id_movimiento='" & rst("id_venta") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            If Val(rst("precio_costo")) <> Val(rstL("costo_promedio")) Then
                X = 0
            End If
        End If
       rst.MoveNext
   Next i
End If




Exit Sub


strCadena = "call put_crear_kardex_general('2018-06-31','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM tmp_kardex_producto1  ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        'SALDO INICIAL
        strCadena = "SELECT ifnull(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v1 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "' "
        Call ConfiguraRstK(strCadena)
        in_ini = rstK(0)
        in_cant_ini = rstK(1)
        'INGRESOS
        strCadena = "SELECT ifnull(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v1 WHERE tipo IN('02','18') and  id_producto='" & rst("id_producto") & "' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
             in_ing = rstK(0)
             in_cant_ing = rstK(1)
        Else
             in_ing = 0
             in_cant_ing = 0
        End If
       
        'SALIDAS
        strCadena = "SELECT ifnull(sum(cantidad_salida*costo_salida),0),ifnull(sum(cantidad_salida),0) FROM kardex_test_v1 WHERE tipo='01' and id_producto='" & rst("id_producto") & "' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
             in_sal = rstK(0)
             in_cant_sal = rstK(1)
        Else
             in_sal = 0
             in_cant_sal = 0
        End If
        
        in_saldo = in_ini + in_ing + in_sal
        
        strCadena = "SELECT ifnull(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v2 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        
        
        
        strCadena = "select `cantidad_saldo_final`(MONTH('2018-04-01'),YEAR('2018-04-01'),k.id_producto,k.id_alm,k.ruc)  as saldo ," & _
        "`precio_promedio_final`(MONTH('2018-04-01'),YEAR('2018-04-01'),k.id_producto,k.id_alm,k.ruc)  as costo " & _
        "from kardex k , producto p where  k.fecha_emision < '2018-04-01' and   k.`id_doc` not in ('0089','0054','0097','0009','0031')  and " & _
        " p.id_linea not in('00017','00009') AND k.`saldo_stock` >=0 and k.id_producto='" & rst("id_producto") & "' and k.id_alm='" & rst("id_alm") & "' and k.ruc='" & KEY_RUC & "' group by k.id_producto, k.id_alm"
        Call ConfiguraRstK(strCadena)
        in_ini = rstK(0)
        in_saldo_stock = rstK(1)
        
        
        
        If Round(in_saldo, 2) <> Round(in_ini, 2) Then
            X = 0
        End If
        rst.MoveNext
        DoEvents
        Me.Command1.Caption = rst("id_producto")
   Next i
End If





Exit Sub
strCadena = "SELECT DISTINCT id_producto,id_alm FROM movimiento_venta v, movimiento_venta_detalle d where v.id_venta=d.id_venta and v.ruc=d.ruc and v.ruc='" & KEY_RUC & "' and v.id_doc='0007' and v.ruc='" & KEY_RUC & "' and v.fecha_emision>='2018-05-01' and v.id_tipo_nota IN ('09') "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
        strCadena = "call put_crear_kardex_producto_v4('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                If j = 0 Then
                        
                    in_saldo = rstK("saldo_stock")
                End If
                in_saldo = in_saldo + rstK("cantidad_real")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                End If
                rstK.MoveNext
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        rst.MoveNext
        DoEvents
        Me.Command1.Caption = str(i) & Space(2) & rst.RecordCount
    Next i
End If

Exit Sub












strCadena = "SELECT * FROM movimiento_venta where id_doc='0007' and ruc='" & KEY_RUC & "' and fecha_emision>='2018-05-01' and id_tipo_nota IN ('09') "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                
                strCadena = "UPDATE  kardex set ruc='46947665' WHERE id_producto='" & rstK("id_producto") & "' and id_movimiento='" & rst("id_venta") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("numero") & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
                CnBd.Execute (strCadena)
                
                strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex WHERE id_producto='" & rstK("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstL(strCadena)
                strCadena = "UPDATE almacen_producto SET stock='" & rstL(0) & "' WHERE id_producto='" & rstK("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
                rstK.MoveNext
            Next j
        End If
        rst.MoveNext
   Next i
End If



Exit Sub







strCadena = "call put_crear_kardex_producto_v3('" & KEY_RUC & "')"
CnBd.Execute (strCadena)
        
strCadena = "SELECT DISTINCT id_producto,id_alm FROM tmp_kardex_producto where  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                End If
               
                rstK.MoveNext
                
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        rst.MoveNext
        DoEvents
        Me.Command1.Caption = str(i) & Space(2) & rst.RecordCount
    Next i
End If
Exit Sub


strCadena = "call put_crear_kardex_producto_v3('" & KEY_RUC & "')"
CnBd.Execute (strCadena)
        
strCadena = "SELECT DISTINCT id_producto,id_alm FROM tmp_kardex_producto where    ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE cantidad>=0 and  id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
               
                
            For j = 0 To rstK.RecordCount - 1
                in_costo_ant = in_costo
                If j = 0 Then
                       in_costo = rstK("costo_promedio")
                Else
                        in_costo = rstK("costo_promedio")
                        If rstK("cantidad_real") < 0 Then
                            If Round(Val(in_costo_ant), 2) <> Round(Val(in_costo), 2) Then
                                strCadena = "UPDATE kardex SET costo_promedio='" & Val(in_costo_ant) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                                CnBd.Execute (strCadena)
                                in_costo = in_costo_ant
                            End If
                        Else
                            If rstK("id_doc") = "0007" Or rstK("id_doc") = "0009" Then
                                If Round(Val(in_costo_ant), 2) <> Round(Val(in_costo), 2) Then
                                strCadena = "UPDATE kardex SET costo_promedio='" & Val(in_costo_ant) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                                CnBd.Execute (strCadena)
                                in_costo = in_costo_ant
                                End If
                            End If
                            
                            If rstK("id_doc") = "0001" Then
                                
                               '  If Round(Val(in_costo_ant), 2) <> Round(Val(in_costo), 2) Then
                               '     strCadena = "UPDATE kardex SET costo_promedio='" & Val(in_costo_ant) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                                '    CnBd.Execute (strCadena)
                                '    in_costo = in_costo_ant
                                'End If
                            End If
                            
                            
                        End If
                    
                End If
                
                
                
                
                
             
                rstK.MoveNext
                DoEvents
            Next j
        End If
        in_costo = 0
        in_costo_ant = 0
        in_saldo = 0
        rst.MoveNext
            DoEvents
        Me.Command1.Caption = str(i) & Space(2) & rst.RecordCount
    Next i
End If
Exit Sub


strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex where   ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
        strCadena = "call put_crear_kardex_producto('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "');"
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                If in_saldo = 0 Then
                    X = 0
                End If
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                End If
               
                rstK.MoveNext
                DoEvents
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        rst.MoveNext
        Me.Command1.Caption = str(i) & Space(2) & rst.RecordCount
    Next i
End If
Exit Sub











Call update_saldo_inicial_vargas(Trim(Me.TxtcodigoProd.Text))
MsgBox "listo"
Exit Sub


GoTo Saldo
strCadena = "SELECT d.id_compra,d.id_producto,d.cantidad,v.numero,v.serie,v.anulado,v.fecha_emision FROM movimiento_compra v,movimiento_compra_detalle d WHERE v.fecha_emision>='2018-01-01' and   v.id_compra=d.id_compra and v.ruc=d.ruc and   d.ruc='" & KEY_RUC & "'  and v.id_doc IN('0001','0003','0050') ORDER BY d.id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM kardex WHERE id_movimiento='" & rst("id_compra") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("numero") & "' and  id_producto='" & rst("id_producto") & "' and cantidad_real='" & rst("cantidad") & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount < 1 Then
            X = 0
        End If
        rst.MoveNext
   Next i
End If



Exit Sub

strCadena = "SELECT d.id_venta,d.id_producto,d.cantidad,v.serie,v.documento,v.anulado,v.fecha_emision FROM movimiento_venta v,movimiento_venta_detalle d WHERE v.fecha_emision>='2018-01-01' and   v.id_venta=d.id_venta and v.ruc=d.ruc and   d.ruc='" & KEY_RUC & "'  and v.id_doc IN('0007') ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM kardex WHERE id_movimiento='" & rst("id_venta") & "' and id_serie='" & rst("serie") & "' and  id_producto='" & rst("id_producto") & "' and cantidad_real='" & rst("cantidad") & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount < 1 Then
            X = 0
        End If
        rst.MoveNext
   Next i
End If



Exit Sub

strCadena = "SELECT sum(cantidad_real),k.`id_producto`,k.id_alm FROM kardex k where k.`fecha_emision`<'2018-02-01' and  ruc='20487725286' group by k.`id_producto`,k.id_alm"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_totalk = 0
   in_totalt = 0
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT cantidad FROM kardex_test  where id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' group by `id_producto`,id_alm"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount < 1 Then
            in_cantidad = 0
        Else
           in_cantidad = rstK(0)
       End If
       in_totalk = in_totalk + rst(0)
       in_totalt = in_totalt + in_cantidad
       If rst(0) <> in_cantidad Then
        X = 0
       End If
       DoEvents
       rst.MoveNext
       Me.Command1.Caption = str(i)
   Next i
End If



Exit Sub
Saldo:







'------- verificar costo unitario
Dim in_ventario(5 To 5)
If KEY_RUC = "20487376338" Then ' casa bebe
    Dim in_stock_actual As Single
    'SELECT id_producto,detalle as nombre_prod,sum(cantidad) FROM view_producto_rotacion WHERE  ruc='20487376338'  GROUP BY id_producto,id_alm  ORDER BY 3 DESC
    'strCadena = "SELECT id_producto,detalle as nombre_prod,sum(cantidad) FROM view_producto_rotacion WHERE  ruc='" & KEY_RUC & "'  GROUP BY id_producto,id_alm  ORDER BY 3 DESC"
    strCadena = "SELECT * FROM producto WHERE id_producto>='02942' and  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
    Call ConfiguraRstChat(strCadena)
    If rstChat.RecordCount > 0 Then
       rstChat.MoveFirst
       For i = 0 To rstChat.RecordCount - 1
           strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and  ruc='" & KEY_RUC & "' ORDER BY id_alm ASC"
           Call ConfiguraRstL(strCadena)
           If rstL.RecordCount > 0 Then
              For j = 0 To rstL.RecordCount - 1
                   strCadena = "SELECT sum(cantidad_real) FROM kardex WHERE id_alm='" & rstL("id_alm") & "' and  id_producto='" & rstChat("id_producto") & "' and ruc='" & KEY_RUC & "'"
                   Call ConfiguraRstT(strCadena)
                   If rstT(0) > 0 Then
                        strCadena = "INSERT INTO inventario_kardex(`id_producto`,`cantidad`,id_alm,`ruc`)VALUES('" & rstChat("id_producto") & "','" & rstT(0) & "','" & rstL("id_alm") & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                        
                   End If
                   rstL.MoveNext
              Next j
           End If
            Call update_kardex_bebe(rstChat("id_producto"))
            
           ' strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' and precio_compra<=0 and precio_venta>0"
           ' Call ConfiguraRstP(strCadena)
           ' If rstP.RecordCount > 0 Then
           '    rstP.MoveFirst
           '    For m = 0 To rstP.RecordCount - 1
           '        strCadena = "SELECT * FROM almacen_producto WHERE precio_compra>0 and   ruc='" & KEY_RUC & "' and id_producto='" & rstP("id_producto") & "' ORDER BY precio_compra DESC LIMIT 1"
           '        Call ConfiguraRstL(strCadena)
           '        If rstL.RecordCount > 0 Then
           '             strCadena = "UPDATE almacen_producto SET precio_compra='" & rstL("precio_compra") & "' WHERE ruc='" & KEY_RUC & "' and id_producto='" & rstP("id_producto") & "' and id_alm='" & rstP("id_alm") & "' LIMIT 1"
           '             CnBd.Execute (strCadena)
           '        End If
           '        rstP.MoveNext
            ''   Next m
           ' End If
            
            
            strCadena = "SELECT * FROM inventario_kardex WHERE id_producto='" & rstChat("id_producto") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
               rstL.MoveFirst
               For k = 0 To rstL.RecordCount - 1
                    Call put_inventario(rstChat("id_producto"), rstL("id_alm"), rstL("cantidad"))
                    rstL.MoveNext
                Next k
            End If
           DoEvents
           
           Me.Command1.Caption = i + 1 & Space(2) + rstChat("id_producto")
           
           rstChat.MoveNext
           DoEvents
          
       Next i
    End If
    MsgBox "se terminoooooo"
    Exit Sub
End If
Exit Sub


If KEY_USUARIO <> "42546269" Then
    Exit Sub
End If

Dim ii As Integer
nuevo:
strCadena = "SELECT * FROM kardex_valorizado_sunat WHERE ruc='" & KEY_RUC & "' and dni_save='" & KEY_USUARIO & "' ORDER BY id_detalle"
Call ConfiguraRstF(strCadena)
If rstF.RecordCount > 0 Then
   rstF.MoveFirst
   ii = 0
   For i = 0 To rstF.RecordCount - 1
        
        If rstF("cantidad_final") = 0 Then
            X = 0
        Else
        in_costo_formula = Round((rstF("cantidad_inicial") * rstF("saldo_inicial") + rstF("saldo_ingreso") - rstF("saldo_salida")) / rstF("cantidad_final"), 2)
        in_costo_promedio = Round(rstF("saldo_final"), 2)
        
        If Val(in_costo_formula) <> Val(in_costo_promedio) Then
            Call update_kardex_VARGAS(rstF("id_producto"))
            DoEvents
            ii = ii + 1
        End If
        End If
        Me.Command1.Caption = str(i + 1) & Space(3) & str(ii) & Space(2) & rstF("id_producto")
       rstF.MoveNext
       DoEvents
   Next i
End If


GoTo nuevo


Exit Sub



Dim in_documento() As String
Dim in_documento2() As String

'*********************** ARREGLAR EL COSTO DE VENTA Y KARDEX
GoTo verificacion

strCadena = "SELECT id_venta,documento,numero_fact FROM movimiento_venta WHERE   anulado='no' and  id_venta>='501427' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_doc in('0007') and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
         
        
        strCadena = "call CON_EliminarVenta('" & rst("id_venta") & "','" & KEY_USUARIO & "') "
        CnBd.Execute (strCadena)
        
        strCadena = "call P_insert_venta_agenda_test('" & rst("id_venta") & "')"
        CnBd.Execute (strCadena)
            
        'MsgBox "NOTA:" + rst("documento") + Space(2) & "REFERENCIA:" & rst("numero_fact"), vbInformation, "MISTER MELLO"
        
        
        rst.MoveNext
        DoEvents
        Me.Command1.Caption = i
        DoEvents
   Next i
End If
MsgBox "PROCESO COMPLETO", vbInformation
Exit Sub

verificacion:
MsgBox "empezamos"
'-VENTAS
Exit Sub

'************************ FIN COSTO VENTA Y KARDEX







' VERIFICACION DE OMPRAS
strCadena = "SELECT * FROM orden_compra WHERE fecha_solicitud>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_solicitud<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY id_compra ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
INICIO2:
        
        strCadena = "SELECT round(sum(cantidad*costo_unitario),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rst("id_doc") & "' and  id_movimiento='" & rst("id_compra") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           in_costo_kardex = rstK("costo_kardex")
        Else
            in_costo_kardex = rstK("costo_kardex")
        End If
        
       
        
        in_costo_conta = 0
        strCadena = "select     (`m`.`DebeMN` + `m`.`HaberMN`) AS `costo`  From     ((`con_documento` `d` join `con_asiento` `a`) join `con_asientomovimiento` `m`) " & _
        "   Where     ((`d`.`Id` = `a`.`IdReferencia`) and (`a`.`Id` = `m`.`IdAsiento`) and (`d`.`IdEmpresaSis` = `a`.`IdEmpresaSis`) and (`a`.`IdEmpresaSis` = `m`.`IdEmpresaSis`) and (`d`.`Activo` = 1) and (`a`.`Activo` = 1) and " & _
        "(`m`.`Activo` = 1) and (`m`.`IdCuentaContable` IN('2501000000000439','2501000000000439')) AND d.IdReferencia='" & rst("id_compra") & "') and d.IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo_conta = rstL("costo")
        Else
            If IsNull(rstK("costo_kardex")) = True And rstL.RecordCount = 0 Then
               ' GoTo sig
            End If
         
             PlaySound App.Path & "\sonidos\dingding.wav"
             'GoTo INICIO
            in_costo_conta = rstL("costo")
        End If
        
        X = rst("documento")
        
        
        If Val(in_costo_conta) <> Val(in_costo_kardex) Then
            If rst("id_doc") = "0007" Then
                strCadena = "call update_costo_venta_vitekey_nota('" & rst("id_venta") & "','" & KEY_RUC & "')"
            Else
                strCadena = "call update_costo_venta_vitekey('" & rst("id_venta") & "','" & KEY_RUC & "')"
            End If
            CnBd.Execute (strCadena)
        End If
        
sig2:
        
        rst.MoveNext
        Me.Command1.Caption = i & Space(2) & rst("fecha_emision")
        DoEvents
   Next i
End If

MsgBox "Proceso completo"
Exit Sub





















strCadena = "SELECT * FROM movimiento_venta WHERE  fecha_emision>='2018-01-01' and fecha_emision<='2018-01-31' and id_doc IN ('0001','0003') and ruc='20128836251'  ORDER BY id_venta DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   '1709 quedo
   
   For i = 0 To rst.RecordCount - 1
        in_detalle = Trim(rst("serie") & rst("numero"))
       strCadena = "SELECT * FROM  costo_venta_empresa WHERE  detalle LIKE '%" & Trim(in_detalle) & "%' "
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount <> 1 Then
           X = 0
       End If
       
       rst.MoveNext
       DoEvents
   Next i
End If

Exit Sub






strCadena = "SELECT * FROM costo_venta_empresa ORDER BY id DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   '1733 quedo
   
   For i = 0 To rst.RecordCount - 1
        in_documento = Split(rst("detalle"), ":")
        in_documento2 = Split(in_documento(1), " ")
        If Len(in_documento2(0)) = 10 Then
            in_serie = Mid(in_documento2(0), 1, 4)
            in_numero = Mid(in_documento2(0), 5, 10)
        Else
           in_serie = Mid(in_documento2(0), 1, 3)
            in_numero = Mid(in_documento2(0), 4, 9)
        End If
        
        strCadena = "SELECT * FROM movimiento_venta WHERE id_doc IN('0003','0001','0008') and  anulado='no' and  numero='" & in_numero & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount <> 1 Then
            X = 0
        End If
        in_venta = rstL("id_venta")
       
       
        strCadena = "SELECT round(sum(cantidad*costo_promedio),4) as costo_kardex FROM kardex WHERE fecha_emision='" & Format(rstL("fecha_emision"), "YYYY-mm-dd") & "' and  id_doc='" & rstL("id_doc") & "' and  id_movimiento='" & rstL("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
           
            in_costo_kardex = rstZ("costo_kardex")
        Else
            in_costo_kardex = rstZ("costo_kardex")
        End If
        
        
        in_costo_conta = 0
    
       If rstZ.RecordCount > 0 Then
            If Round(in_costo_kardex, 4) <> Round(rst("costo"), 4) Then
                strCadena = "call update_costo_venta_vitekey('" & in_venta & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            Else
                strCadena = "UPDATE costo_venta_empresa SET procesado='1' WHERE id='" & rst("id") & "'"
                CnBd.Execute (strCadena)
            End If
       Else
            X = 0
       End If
       rst.MoveNext
       Me.Command1.Caption = str(i) & Space(2) & str(rst.RecordCount)
       DoEvents
   Next i
End If
Exit Sub


'*********************** compras
strCadena = "SELECT * FROM movimiento_compra WHERE fecha_emision>='2018-01-01' and   id_doc='0001' and  ruc='" & KEY_RUC & "' ORDER BY id_compra ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   For i = 0 To rst.RecordCount - 1
   

            strCadena = "SELECT  sum(am.DebeMN),sum(am.HaberMN),IdAsiento From `con_documento` d, `con_asiento` a, `con_asientomovimiento` am " & _
            " Where d.`Id`=a.`IdReferencia` and a.`Id`=am.`IdAsiento` and d.`IdEmpresaSis`=a.`IdEmpresaSis` and d.`Activo`=1 and a.`Activo`=1 and " & _
            " am.`Activo`=1 and  d.`IdReferencia`='" & rst("id_compra") & "' and d.IdEmpresaSis='" & KEY_RUC & "'"
            Call ConfiguraRstK(strCadena)
            in_acum1 = rstK(0)
            in_acum2 = rstK(1)
            If IsNull(rstK(0)) = True Then
                GoTo s
            End If
            If Val(in_acum1) <> Val(in_acum2) Then
                 strCadena = "UPDATE con_asientomovimiento set DebeMN=round(DebeMN,2),HaberMN=round(HaberMN,2) WHERE IdAsiento='" & rstK("IdAsiento") & "' and IdEmpresaSis='" & KEY_RUC & "' "
                 CnBd.Execute (strCadena)
                
            End If
s:

            rst.MoveNext
            DoEvents
   Next i
   
End If
'*******************************

Exit Sub







ini:
strCadena = "SELECT  sum(cantidad*costo_unitario), id_movimiento,id_producto,id_kardex FROM kardex WHERE id_doc not in('0009','0090','0089','0106') and  ruc='" & KEY_RUC & "' and cantidad_real<0 and fecha_emision<='2018-04-31' group by id_movimiento,id_producto ORDER BY id_kardex DESC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT sum(cantidad*precio_costo) from movimiento_venta_detalle WHERE id_venta='" & rst("id_movimiento") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' GROUP BY id_producto "
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
       If Round(rst(0), 2) <> Round(rstK(0), 2) Then
            X = rst("id_movimiento")
            Y = X
           ' GoTo ini
       End If
       Else
            Y = rst("id_movimiento")
            Y = 1
          '  GoTo ini
            
       End If
       rst.MoveNext
       DoEvents
   Next i
End If

Exit Sub


strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' and fecha_emision>='2018-01-01' and fecha_emision<='2018-01-31' and id_doc IN('0001','0003','0007','0008')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT * FROM view_costo_venta_venta WHERE id_venta='" & rst("id_venta") & "'"
   Next i
   
End If








strCadena = "SELECT v.`id_venta`,v.`documento`,v.`numero`FROM movimiento_venta  v Where fecha_emision>='2018-01-01' and fecha_emision<='2018-01-31' and " & _
" id_doc  IN ('0007') and anulado='no' and  v.`id_doc`  and ruc='20128836251'  and v.`id_tipo_nota` NOT IN ('04','05') order by fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
         X = rst("documento")
        
     Y = rst("id_venta")
  '  strCadena = "call CON_EliminarVenta('" & rst("id_venta") & "','" & KEY_USUARIO & "') "
 'CnBd.Execute (strCadena)
     
     
  '  strCadena = "call P_insert_venta_agenda_test('" & rst("id_venta") & "')"
  '  CnBd.Execute (strCadena)
       
  ' strCadena = "call update_costo_venta_vargas('" & rst("id_venta") & "')"
  '  CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT ca.`Id` FROM movimiento_venta vm,con_documento cd ,con_asiento ca where " & _
    " vm.`id_venta`=cd.`IdReferencia` and cd.`Id`=ca.`IdReferencia` and vm.`ruc`=cd.`IdEmpresaSis` and cd.`Activo`=1 and ca.Activo=1 and vm.`id_venta`='" & rst("id_venta") & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        in_asiento = rstL(0)
    
    
    strCadena = "select sum(DebeMN),sum(HaberMN) from con_asientomovimiento m where m.`IdAsiento`='" & in_asiento & "'"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        If rstK(0) <> rstK(1) Then
           strCadena = "DELETE FROM con_asientomovimiento WHERE `IdAsiento`='" & in_asiento & "' ORDER BY id DESC LIMIT 1"
           CnBd.Execute (strCadena)
        End If
    End If
    End If
    
        
    rst.MoveNext
    DoEvents
   Next i
End If

Exit Sub








Dim in_producto As String
strCadena = "SELECT * FROM `inventario_vargas_31122017"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
   
   in_producto = Format(rst("id_producto"), "00000")
   strCadena = "SELECT id_kardex FROM kardex WHERE id_doc='0106' and cantidad='" & rst("cantidad") & "' and  fecha_emision='2017-12-31'  and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount < 1 Then
        Call update_saldo_inicial_vargas(Trim(Me.TxtcodigoProd.Text))
  Else
        If rstK.RecordCount > 1 Then
            X = 0
        End If
   End If
   rst.MoveNext
   Me.Command1.Caption = i
   'DoEvents
   Next i
End If
MsgBox "ya"

Exit Sub





strCadena = "select p.id_producto,p.`nombre_prod` from producto p, tipo_producto t Where p.`id_tipo`=t.`id_tipoproducto` and p.`ruc`=t.`ruc` and t.servicio='si' and  p.ruc='20128836251'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "DELETE FROM kardex WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
    Next i
End If





strCadena = "select v.`documento`,d.`precio_costo`,d.`id_producto`,v.id_venta,v.fecha_emision from movimiento_venta v,movimiento_venta_detalle D Where " & _
"d.`precio_costo`<=0 and v.`ruc`='" & KEY_RUC & "' and v.id_venta=d.`id_venta` and v.`ruc`=d.`ruc` and " & _
" v.`id_doc` IN ('0001','0003','0007','0008') order by d.`id_detalle_venta`"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "SELECT costo_promedio FROM kardex WHERE costo_promedio>0 and  fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_movimiento='" & rst("id_venta") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            strCadena = "UPDATE movimiento_venta_detalle SET precio_costo='" & rstK("costo_promedio") & "' WHERE id_venta='" & rst("id_venta") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        rst.MoveNext
   Next i
End If



strCadena = "SELECT * FROM kardex WHERE id_doc IN('0001','0003') and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
End If





strCadena = "call put_update_costo_venta_kardex('" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM tmp_kardex2 "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "UPDATE movimiento_venta_detalle SET precio_costo='" & rst("costo_promedio") & "'  WHERE cantidad='" & rst("cantidad") & "' and  id_producto='" & rst("id_producto") & "' and id_venta='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       CnBd.Execute (strCadena)
       rst.MoveNext
       DoEvents
   Next i
End If



Exit Sub
strCadena = "SELECT * FROM inventario_vargas_31122017   "
Call ConfiguraTemporal(strCadena)
If rstTemporal.RecordCount > 0 Then
   rstTemporal.MoveFirst
   For i = 0 To rstTemporal.RecordCount - 1
       strCadena = "SELECT * FROM kardex WHERE id_producto='" & Format(rstTemporal("id_producto"), "00000") & "' and id_doc='0106' and fecha_emision='2017-12-31'"
       Call ConfiguraRstlocal(strCadena)
       If rstLocal.RecordCount < 1 Then
          Call update_kardex_VARGAS(Format(rstTemporal("id_producto"), "00000"))
       End If
       rstTemporal.MoveNext
       Me.Command1.Caption = str(i + 1) & Space(2) & Space(1) & rstTemporal.RecordCount
       DoEvents
       
   Next i
End If

End Sub





Private Sub DtcProducto_KeyPress(KeyAscii As Integer)
On Error GoTo salir
 If KeyAscii = 13 Then
    Call presionar
End If
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub presionar()
'If KEY_BARRAS = "si" Then
    strCadena = "SELECT cod_barra FROM producto_barras WHERE id_producto='" & Trim(Me.DtcProducto.BoundText) & "' AND ruc='" & KEY_RUC & "'"
'Else
strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & Trim(Me.DtcProducto.BoundText) & "' AND ruc='" & KEY_RUC & "'"
'End If

 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Me.TxtcodigoProd.Text = rst(0)
    Call Resalta(Me.TxtcodigoProd)
    
 End If
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 0
strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  Me.TxtcodigoProd.Enabled = False

  
strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
Me.DtcProducto.Enabled = False

If KEY_USUARIO = "42546269" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "900001" Then
    Me.cmdUpdateKardexTotal.Visible = True
    Me.Command1.Visible = True
Else
    Me.cmdUpdateKardexTotal.Visible = False
    Me.Command1.Visible = False
End If

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_EXIT
        Unload Me
End Select
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TxtBusquedaRapida_Change()
If Len(Me.TxtBusquedarapida) > 0 Then
    strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' AND nombre_prod LIKE '%" & Trim(Me.TxtBusquedarapida.Text) & "%' ORDER BY nombre_prod"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProducto)
End If
End Sub

Private Sub TxtBusquedarapida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcProducto.Enabled = True Then
        Me.DtcProducto.SetFocus
    End If
End If
End Sub


Private Sub load_kardex()
Dim stcodigop As String
Dim stock_anterior As Double, stock As Double
Dim fecha_inicio As Date
Dim Campo As Field
Dim arrColWidth()   As Long
stock_anterior = 0
stock = 0
Call LlenarKardexDetallado




End Sub
Private Sub TxtcodigoProd_KeyPress(KeyAscii As Integer)
On Error GoTo salir
If KeyAscii = 13 Then
   Call load_kardex
Exit Sub

salir:
MsgBox "RECUERDA.... HACER SOLO [1] ENTER", vbInformation
Exit Sub





































If KEY_BARRAS = "si" Then
    strCadena = "SELECT id_producto FROM producto_barras  WHERE cod_barra='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "' "
Else
   Me.TxtcodigoProd.Text = formato_item(Me.TxtcodigoProd.Text, 5)
    strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "' "
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 And Me.TxtcodigoProd.Text <> "" Then
    stcodigop = rst(0)
Else
   Set rst = Nothing
   Procedencia = buscar
   FrmProducto.Show
   Exit Sub
End If

If Me.chkBuscarfechas.Value = 0 Then
    strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' "
Else
    strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' "
End If
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = False Then
    stock_anterior = rst(0)
Else
    stock_anterior = 0
    Me.HfdGrilla.Rows = 0
    Me.HfdGrilla.Clear
End If


Set rst = Nothing
strCadena = "SELECT P.stock_minimo,U.abreviatura FROM producto P,unidad U WHERE id_producto='" & stcodigop & "' AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)


Set rst = Nothing
strCadena = "SELECT fecha_emision,CONCAT(C.doc_abrev,':',K.id_serie,'-',K.id_numero) as comprobante,cantidad_ing,cantidad_sal,cantidad_real,precio,ncliente FROM kardex K,comprobantes C  WHERE K.id_doc=C.id_doc AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "' ORDER BY K.id_kardex DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            Exit Sub
        End If
        Me.DtcProducto.BoundText = Trim(stcodigop)
        N = 1
       Me.HfdGrilla.Clear
       Me.HfdGrilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Me.HfdGrilla.ColWidth(0) = 1150
            Me.HfdGrilla.ColWidth(1) = 1850
            Me.HfdGrilla.ColWidth(2) = 4500
            Me.HfdGrilla.ColWidth(3) = 1300
            Me.HfdGrilla.ColWidth(4) = 1300
            Me.HfdGrilla.ColWidth(5) = 1300
            Me.HfdGrilla.ColWidth(6) = 1400
        Next
        
       cabecera = "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "INGRESOS" & vbTab & "SALIDAS" & vbTab & "STOCK FINAL"
       Me.HfdGrilla.AddItem cabecera
         For k = 0 To 6
             HfdGrilla.col = k
             HfdGrilla.Row = 0
             HfdGrilla.CellBackColor = &HDFDFE0
        Next k
     '   Fila = "***********" & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & rst("cantidad_ing") & vbTab & rst("cantidad_sal") & vbTab & Format(stock_anterior, "#,##0.00")
      '  Me.HfdGrilla.AddItem Fila
            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("fecha_emision") & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & rst("cantidad_ing") & vbTab & rst("cantidad_sal") & vbTab & Format(stock_anterior, "#,##0.00")
            Me.HfdGrilla.AddItem Fila
            If rst("cantidad_real") > 0 Then
                stock_anterior = stock_anterior + rst("cantidad_real")
            Else
                stock_anterior = stock_anterior - rst("cantidad_real")
            End If
                    For j = i To Me.HfdGrilla.Rows - 1
                        Me.HfdGrilla.col = 5
                        Me.HfdGrilla.Row = j
                        Me.HfdGrilla.CellBackColor = &HC0FFC0
                    Next j
            Fila = ""
            rst.MoveNext
        Next i
        
       ' Call LlenaDescripcion(Trim(stcodigop), Trim(Me.DtcAlmacen.BoundText))
        'Call Resalta(Me.TxtcodigoProd)
    End If
    
End Sub




Private Sub LlenarKardexDetallado()

Iniciar:
Dim in_costo_anterior As Double
Dim in_costo_promedio_ant As Double

Dim stcodigop As String
Dim stock_anterior As Double, stock As Double, stock_inicial As Double
Dim fecha_inicio As Date, precio_costo As Double, Monto1 As Double, Monto2 As Double, valorizado As Double
Dim Campo As Field
Dim arrColWidth()   As Long
stock_anterior = 0
stock = 0
stock_inicial = 0

If KEY_BARRAS = "si" Then
    Me.TxtcodigoProd.Text = Format(Me.TxtcodigoProd.Text, "00000")
    strCadena = "SELECT id_producto FROM producto_barras  WHERE cod_barra='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "' "
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount < 1 Then
       strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "' "
    End If
Else
    Me.TxtcodigoProd.Text = Format(Me.TxtcodigoProd.Text, "00000")
    strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' AND ruc='" & KEY_RUC & "' "
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 And Me.TxtcodigoProd.Text <> "" Then
    stcodigop = rst(0)
Else
   Set rst = Nothing
   Procedencia = buscar
   FrmProducto.Show
   Exit Sub
End If

If Me.chkBuscarfechas.Value = 0 Then
    If Me.chk_all.Value = 1 Then
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "' AND  ruc='" & KEY_RUC & "' "
        Else
            strCadena = "SELECT SUM(cantidad_factura) AS Expr1 FROM   kardex WHERE     id_producto='" & Trim(stcodigop) & "' AND  ruc='" & KEY_RUC & "' "
        End If
    Else
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' "
        Else
            strCadena = "SELECT SUM(cantidad_factura) AS Expr1 FROM   kardex WHERE   id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' "
        End If
    End If
    stock_inicial = 0
Else
    If Me.chk_all.Value = 1 Then
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "'  AND ruc='" & KEY_RUC & "' AND fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "'"
        Else
            strCadena = "SELECT SUM(cantidad_factura) AS Expr1 FROM   kardex WHERE  id_producto='" & Trim(stcodigop) & "'  AND ruc='" & KEY_RUC & "' AND fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "'"
        End If
    Else
         If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT SUM(cantidad_real) AS Expr1 FROM   kardex WHERE id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "'"
         Else
            strCadena = "SELECT SUM(cantidad_factura) AS Expr1 FROM   kardex WHERE  id_producto='" & Trim(stcodigop) & "' AND id_alm ='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "'"
         End If
    End If
    
    
End If

Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = False Then
    stock_anterior = rst(0)
    If Me.chkBuscarfechas.Value = 0 Then
       stock_inicial = 0
    Else
        stock_inicial = rst(0)
    End If
Else
    stock_anterior = 0
    stock_inicial = 0
    Me.HfdGrilla.Rows = 0
   
End If
If Me.chkBuscarfechas.Value = 0 Then
    'strCadena = "UPDATE almacen_producto SET stock ='" & Val(stock_anterior) & "' WHERE id_producto='" & Trim(stcodigop) & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
End If



strCadena = "SELECT P.stock_minimo,U.abreviatura FROM producto P,unidad U WHERE id_producto='" & stcodigop & "' AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)



If Me.chkBuscarfechas.Value = 0 Then
    If Me.chk_all.Value = 1 Then
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT * FROM view_kardex WHERE  id_tipo_movimiento<>'10' and id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT * FROM view_kardex WHERE cantidad_factura<>0 and  id_doc IN('0001','0003','0007','0009','0089,'0090') and  id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "'"
        End If
    Else
         If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT * FROM view_kardex WHERE id_tipo_movimiento<>'10' and id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "'"
         Else
            strCadena = "SELECT * FROM view_kardex WHERE cantidad_factura<>0 and id_doc IN('0001','0003','0007','0009','0089','0090') and id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "'"
         End If
    End If
Else
    If Me.chk_all.Value = 1 Then
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT * FROM view_kardex WHERE  id_tipo_movimiento<>'10' and id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
        Else
            strCadena = "SELECT * FROM view_kardex WHERE cantidad_factura<>0 and id_doc IN('0001','0003','0007','0009','0089','0090') and id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
        End If
    Else
        If Me.chk_kardex_fisico.Value = 1 Then
            strCadena = "SELECT * FROM view_kardex WHERE id_tipo_movimiento<>'10' and  id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
        Else
            strCadena = "SELECT * FROM view_kardex WHERE cantidad_factura<>0 and id_doc IN('0001','0003','0007','0009','0089','0090') and id_alm='" & Me.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(stcodigop) & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
        End If
    End If
    
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   Me.HfdGrilla.Rows = 0
   Exit Sub
        End If
       Me.DtcProducto.BoundText = Trim(stcodigop)
       
       Me.HfdGrilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Me.HfdGrilla.ColWidth(0) = 1200
            Me.HfdGrilla.ColWidth(1) = 2000
            Me.HfdGrilla.ColWidth(2) = 3800
            Me.HfdGrilla.ColWidth(3) = 1100
            Me.HfdGrilla.ColWidth(4) = 1500
            Me.HfdGrilla.ColWidth(5) = 1500
            Me.HfdGrilla.ColWidth(6) = 1200
            Me.HfdGrilla.ColWidth(7) = 1500
            Me.HfdGrilla.ColWidth(8) = 1500
            Me.HfdGrilla.ColWidth(9) = 1200
            Me.HfdGrilla.ColWidth(10) = 1500
            Me.HfdGrilla.ColWidth(11) = 1500
        Next
        
       cabecera = "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "CANTIDAD" & vbTab & "COSTO.U" & vbTab & "TOTAL " & vbTab & "CANTIDAD" & vbTab & "COSTO.U" & vbTab & "TOTAL " & vbTab & "CANTIDAD" & vbTab & "COSTO.U" & vbTab & "TOTAL "
       Me.HfdGrilla.AddItem cabecera
         For k = 0 To 11
             HfdGrilla.col = k
             HfdGrilla.Row = 0
             HfdGrilla.CellBackColor = &HDFDFE0
        Next k
        
        
        
        Fila = " " & vbTab & "INVENTARIO INICIAL" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ValorItemII(preciosal * rst("cantidad_sal")) & vbTab & Format(stock_inicial, "#,##0.00") & vbTab & ValorItemII(rst("costo_unitario")) & vbTab & ValorItemII(stock_inicial * rst("costo_unitario"))
        Me.HfdGrilla.AddItem Fila
            
            
        
        rst.MoveFirst
        cant_real = 0
        
        in_cantidad_ingreso = 0
        in_saldo_ingresos = 0
        in_cantidad_salidas = 0
        in_saldo_salidas = 0
        in_cantidad_saldo = 0
        in_saldo_final = 0
        For i = 0 To rst.RecordCount - 1
            
            cant_real = rst("cantidad_real") '+ rst("cantidad_contable")
            If cant_real > 0 Then
                   precioing = rst("costo_unitario")
                   preciosal = 0
                   Monto1 = precioing * rst("cantidad_ing")
                   
            Else
                   Monto1 = rst("costo_unitario") * cant_real
                   precioing = 0
                   preciosal = rst("costo_unitario")
                   If Me.chk_kardex_contable.Value = 1 And rst("id_tipo_movimiento") = "10" Then
                        cant_real = rst("cantidad_factura")
                   End If
            End If
            
            
            stock_inicial = stock_inicial + cant_real
            
            
            costo_movil = rst("costo_promedio")
            valorizado = stock_inicial * costo_movil
            
            
            Fila = Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & ValorItem(rst("cantidad_ing")) & vbTab & ValorItemII(precioing) & vbTab & ValorItem(precioing * rst("cantidad_ing")) & vbTab & ValorItem(rst("cantidad_sal")) & vbTab & ValorItemII(preciosal) & vbTab & ValorItem(preciosal * rst("cantidad_sal")) & vbTab & ValorItem_v(stock_inicial) & vbTab & ValorItemIII(costo_movil) & vbTab & ValorItem_v(valorizado)
            Me.HfdGrilla.AddItem Fila
            DoEvents
            
            
            '***************************************************Aux Modificacion Costos
            If Me.chk_costo_promedio.Value = 1 And (KEY_USUARIO = "42546269" Or KEY_USUARIO = "46947665") And rst("id_doc") <> "0106" Then
                'salidas
                    If rst("cantidad_real") < 0 And i > 0 Then

                        rst.MovePrevious
                        in_costo_anterior = rst("costo_promedio")
                        in_costo_promedio_ant = rst("costo_promedio")
                        in_valorizado_anterior = rst("saldo_stock") * rst("costo_promedio")
                        rst.MoveNext
                        
                        If Val(stock_inicial) <> rst("saldo_stock") Then
                            strCadena = "UPDATE kardex SET saldo_stock='" & Val(stock_inicial) & "'  WHERE id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                            CnBd.Execute (strCadena)
                           GoTo Iniciar
                        End If
                        DoEvents
                    Else
                     
                        
                        If rst("id_doc") = "0009" Or rst("id_doc") = "0031" Then
                            
                        Else
                            in_costo_anterior = rst("costo_unitario")
                        End If
                        
                        
                            rst.MovePrevious
                            If rst.BOF Then
                                in_valorizado_anterior = 0
                                rst.MoveNext
                                 DoEvents
                            Else
                                in_valorizado_anterior = rst("saldo_stock") * rst("costo_promedio")
                                rst.MoveNext
                                 DoEvents
                            End If
                            
                     
                        
                    End If
                    
                    If rst("cantidad_real") < 0 Then
                        If Abs(rst("costo_unitario") - Val(in_costo_anterior)) > 0.0001 Then
                           strCadena = "UPDATE kardex SET costo_unitario='" & in_costo_anterior & "'  WHERE id_kardex='" & rst("id_kardex") & "' and id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                           CnBd.Execute (strCadena)
                           GoTo Iniciar
                           Exit Sub
                        End If
                        
                    Else
                        
                        If rst("id_doc") = "0009" Or rst("id_doc") = "0031" Then
                            in_costo_anterior = get_costo_unitario(rst("id_producto"), rst("id_doc"), rst("id_serie"), rst("id_numero"))
                            If rst("costo_unitario") <> in_costo_anterior Then
                               strCadena = "UPDATE kardex SET costo_unitario='" & in_costo_anterior & "'  WHERE id_kardex='" & rst("id_kardex") & "' and id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                               CnBd.Execute (strCadena)
                               GoTo Iniciar
                            End If
                        Else
                            If rst("id_doc") = "0007" Then
                                in_costo_anterior = get_costo_unitario_nota(rst("id_producto"), rst("id_alm"), rst("id_kardex"))
                                If rst("costo_unitario") <> in_costo_anterior Then
                                    strCadena = "UPDATE kardex SET costo_unitario='" & in_costo_anterior & "'  WHERE id_kardex='" & rst("id_kardex") & "' and id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                                    CnBd.Execute (strCadena)
                                    GoTo Iniciar
                                End If
                            End If
                        End If
                        
                    End If
                    
                    If rst("cantidad_real") < 0 Then
                        in_valorizado_actual = rst("cantidad_real") * in_costo_anterior
                    Else
                        in_valorizado_actual = rst("cantidad_real") * rst("costo_unitario")
                    End If
                    
                    If stock_inicial = 0 Then
                        in_costo_promedio = in_costo_anterior
                        
                    Else
                        in_costo_promedio = Val((in_valorizado_anterior + in_valorizado_actual)) / stock_inicial
                    End If
                    
                    
                    If rst("cantidad_real") < 0 Then
                       If Abs(rst("costo_promedio") - Val(in_costo_promedio)) <> 0 Then
                           ' Call put_costo_promedio(rst("id_kardex"), rst("id_producto"), Val(in_costo_promedio), rst("cantidad_real"))
                            GoTo Iniciar
                       End If
                    Else
                    
                    End If
                    
            End If
            
            stock_anterior = stock_anterior - cant_real
            
            in_cantidad_ingreso = in_cantidad_ingreso + rst("cantidad_ing")
            in_saldo_ingresos = in_saldo_ingresos + precioing * rst("cantidad_ing")
            
            in_cantidad_salidas = in_cantidad_salidas + rst("cantidad_sal")
            in_saldo_salidas = in_saldo_salidas + preciosal * rst("cantidad_sal")
            
            in_cantidad_saldo = in_cantidad_saldo + stock_inicial
            in_saldo_final = in_saldo_final + valorizado
           
           
          ' For k = 9 To 11
          '                      HfdGrilla.col = k
          '                      HfdGrilla.Row = i + 1
          '                      HfdGrilla.CellBackColor = &HC0FFC0
          'Next k
                            
              
           ' DoEvents
            rst.MoveNext
        Next i
        
        
            
        For k = 9 To 11
                                HfdGrilla.col = k
                                HfdGrilla.Row = i + 1
                                HfdGrilla.CellBackColor = &H8080FF
        Next k
        


End Sub




Private Sub put_costo_promedio_almacen(ByVal in_producto As String, ByVal in_alm As String)

    strCadena = "SELECT costo_promedio FROM kardex WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1"
    Call ConfiguraRstChat2(strCadena)
    If rstchat2.RecordCount > 0 Then
       If rstchat2("costo_promedio") > 0.01 Then
           strCadena = "UPDATE almacen_producto SET precio_compra='" & rstchat2("costo_promedio") & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
       End If
    End If

   



End Sub
Private Function get_costo_unitario(ByVal in_producto As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Double
strCadena = "SELECT costo_unitario FROM kardex WHERE id_doc='" & in_doc & "' and   cantidad_real<0 and  id_producto='" & in_producto & "' and id_serie='" & in_serie & "' and id_numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    get_costo_unitario = rstT("costo_unitario")
Else
    get_costo_unitario = 0
End If
End Function

Private Function get_costo_unitario_nota(ByVal in_producto As String, ByVal in_alm As String, ByVal in_kardex As String) As Double
strCadena = "SELECT costo_unitario FROM kardex WHERE  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' and id_kardex<'" & Val(in_kardex) & "' ORDER BY id_kardex DESC LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    get_costo_unitario_nota = rstT("costo_unitario")
Else
    get_costo_unitario_nota = 0
End If
End Function



Private Sub LlenaDescripcion(ByVal codigo As String, ByVal Almacen As String)
strCadena = "SELECT Producto.DescripcionProducto,Unidad.sAbreviatura,Producto.StockMinimo FROM Producto INNER JOIN Unidad ON Producto.cUnidad=Unidad.cUnidad WHERE Producto.cProducto='" & codigo & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
   
    strCadena = "SELECT     SUM(Stk_Cant) AS Expr1 From Kardex WHERE  cProducto = '" & Trim(codigo) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
    Call ConfiguraRst(strCadena)
On Error GoTo 100
     '   Me.TxtStockActual.Text = str(rst(0)) + Space(32) + Me.txtunidad.Text
     '   Set rst = Nothing
        Exit Sub
100:
      '  Me.TxtStockActual.Text = "0" + Space(32) + Me.txtunidad.Text

    Set rst = Nothing
End If
End Sub
Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub



VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18795
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command30 
      Caption         =   "GENERAR UN WORD"
      Height          =   375
      Left            =   13320
      TabIndex        =   67
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdLenin 
      Caption         =   "MIGRAR LENIN OLIVOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   66
      Top             =   7800
      Width           =   3375
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H000080FF&
      Caption         =   "SALDOS INICIAL"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7080
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8400
      OleObjectBlob   =   "frmpanel.frx":0000
      Top             =   1440
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H000080FF&
      Caption         =   "SALDOS"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox txtAlmacen3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   60
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox txtAlmacen2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   59
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox txtAlmacen1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   58
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command29 
      Caption         =   "PLAN CONTABLE"
      Height          =   495
      Left            =   6000
      TabIndex        =   57
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H000080FF&
      Caption         =   "UPDATE SELVA  V2"
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6360
      Width           =   3495
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPDATE SELVA V1"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6000
      Width           =   3495
   End
   Begin VB.CommandButton Command24 
      Caption         =   "CUENTAS COBRAR CONTADO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   50
      Top             =   9600
      Width           =   3735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "RESTAURAR COMPRAS"
      Height          =   375
      Left            =   12960
      TabIndex        =   49
      Top             =   1320
      Width           =   3180
   End
   Begin VB.CommandButton Command22 
      Caption         =   "RESTAURAR DANIEL OLIVOS"
      Height          =   375
      Left            =   12960
      TabIndex        =   48
      Top             =   960
      Width           =   3180
   End
   Begin VB.CommandButton Command21 
      Caption         =   "CARGAR CUENTA PAGAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      TabIndex        =   47
      Top             =   6840
      Width           =   3255
   End
   Begin VB.CommandButton Command20 
      Caption         =   "MIGRAR  COMPRAS VARGAS"
      Height          =   375
      Left            =   13440
      TabIndex        =   46
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "MIGRAR KARDEX VARGAS"
      Height          =   1695
      Left            =   9720
      TabIndex        =   45
      Top             =   6000
      Width           =   7215
      Begin VB.CommandButton Command26 
         BackColor       =   &H008080FF&
         Caption         =   "GENERAR KARDEX VALORIZADO"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   840
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   166330369
         CurrentDate     =   43322
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   375
         Left            =   1680
         TabIndex        =   52
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   166330369
         CurrentDate     =   43322
      End
   End
   Begin VB.CommandButton cmdCuentaCobrar 
      Caption         =   "MIGRAR CUENTAS POR COBRAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   44
      Top             =   9120
      Width           =   3735
   End
   Begin VB.CommandButton cmdleerExcelCuentasCobrar 
      Caption         =   "LEER EXCEL FORMATO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   42
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "KARDEX CHACHA"
      Height          =   615
      Left            =   4800
      TabIndex        =   41
      Top             =   5280
      Width           =   3495
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   4920
      Top             =   9240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   375
      Left            =   13320
      TabIndex        =   40
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "ACTUALIZAR STOCK PRECIO"
      Height          =   375
      Left            =   9600
      TabIndex        =   39
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "LEER EXCEL FORMATO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   37
      Top             =   1320
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   4800
      Top             =   4680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command16 
      Caption         =   "MIGRAR ALUMNOS BIBLIOTECA"
      Height          =   375
      Left            =   9600
      TabIndex        =   36
      Top             =   960
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4680
      TabIndex        =   34
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdcompras 
      Caption         =   "MIGRAR  COMPRAS"
      Height          =   375
      Left            =   4680
      TabIndex        =   33
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton cmdmigrar_clientes 
      Caption         =   "MIGRAR CLIENTES"
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmdventas 
      Caption         =   "MIGRAR VENTAS"
      Height          =   375
      Left            =   4680
      TabIndex        =   31
      Top             =   3480
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4920
      Top             =   9120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.OptionButton OptAccess 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ACCESS"
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   16800
      TabIndex        =   30
      Top             =   1600
      Width           =   2895
   End
   Begin VB.CommandButton cmdtest 
      BackColor       =   &H008080FF&
      Caption         =   "REALIZAR CONEXION"
      Height          =   375
      Left            =   16800
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2040
      Width           =   2895
   End
   Begin VB.OptionButton optsql2005 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SQL>2005"
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   16800
      TabIndex        =   28
      Top             =   1260
      Width           =   2895
   End
   Begin VB.OptionButton optsql200 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SQL 2000"
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   16800
      TabIndex        =   27
      Top             =   930
      Width           =   2895
   End
   Begin VB.OptionButton optpostgres 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "POSTGRES"
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   16800
      TabIndex        =   26
      Top             =   580
      Width           =   2895
   End
   Begin VB.OptionButton optmysql 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "MYSQL"
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   16800
      TabIndex        =   25
      Top             =   240
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.CommandButton cmdunidades 
      Caption         =   "INICIAR MIGRACION"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "MIGRAR PRECIO COSTO"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   9720
      Width           =   3660
   End
   Begin VB.CommandButton Command12 
      Caption         =   "MIGRAR PRODUCTOS FRACCIONADOS"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   9240
      Width           =   3660
   End
   Begin VB.CommandButton Command11 
      Caption         =   "MIGRAR EMPLEADORAS-SEGUROS"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   8760
      Width           =   3660
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ACTUALIZAR_SUBLINEAS"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   8280
      Width           =   3660
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ELIMINAR ACTIVIDADES"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   7800
      Width           =   3660
   End
   Begin VB.CommandButton cmdactuallizar_insumos 
      Caption         =   "MATERIAL QUIRURJICO - INSUMOS"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   7320
      Width           =   3660
   End
   Begin VB.CommandButton Command9 
      Caption         =   "MIGRAR PRECIOS FARMACIA"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   6840
      Width           =   3660
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ELIMINARPRESION ARTERIOR"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6360
      Width           =   3660
   End
   Begin VB.CommandButton Command7 
      Caption         =   "MIGRAR FARMACIA"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   5880
      Width           =   3660
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MIGRAR AMBULANCIA"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5400
      Width           =   3660
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ELIMINAR AGENDA "
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   3660
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ELIMINAR PACIENTES DEL HRL"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   3660
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACTUALIZAR SERVICIOS"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   3660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MODIFICAR RUC"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3660
   End
   Begin VB.CommandButton cmdmigracionservicios 
      Caption         =   "GRUPO CONTABLE"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   3660
   End
   Begin VB.TextBox txtusuario 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "42546269"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdlinea 
      Caption         =   "MIGRACION  ASEGURADORA"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3660
   End
   Begin VB.TextBox txtruc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "20487725286"
      Top             =   600
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfproductos 
      Height          =   1215
      Left            =   9600
      TabIndex        =   38
      Top             =   1800
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2143
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuentasCobrar 
      Height          =   1695
      Left            =   9600
      TabIndex        =   43
      Top             =   4080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2990
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALM 3:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6120
      TabIndex        =   63
      Top             =   8520
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALM 2:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6120
      TabIndex        =   62
      Top             =   8160
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALM 1:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   6120
      TabIndex        =   61
      Top             =   7800
      Width           =   510
   End
   Begin VB.Label lbldireccion 
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      TabIndex        =   54
      Top             =   1440
      Width           =   3765
   End
   Begin VB.Label lblItem 
      Caption         =   " "
      Height          =   255
      Left            =   5640
      TabIndex        =   35
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblempresa 
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      TabIndex        =   24
      Top             =   1080
      Width           =   3765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PANEL DE MIGRACION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   8400
      TabIndex        =   22
      Top             =   120
      Width           =   4485
   End
   Begin VB.Label lblcantidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1320
      TabIndex        =   15
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO :"
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
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC :"
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
      Left            =   1230
      TabIndex        =   1
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS :"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   14505
      Left            =   -2160
      Picture         =   "frmpanel.frx":0234
      Stretch         =   -1  'True
      Top             =   -2040
      Width           =   28800
   End
End
Attribute VB_Name = "frmpanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Click()
strCadena = "SELECT * FROM producto WHERE id_tipo='07' and ruc='20487911586'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("presentacion_und") = "-" Then
          ' strCadena = "UPDATE producto SET presentacion ='" & get_numero(rst("presentacion")) & "',presentacion_und='" & get_unidad(rst("presentacion")) & "' WHERE id_producto='" & rst("id_producto") & "' and ruc='20487911586'"
          ' CnBd.Execute (strCadena)
       End If
       rst.MoveNext
   Next i
End If

End Sub
Public Sub migrar_marca()

strCadena = "SELECT * FROM ginsac_marca order by id"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO marca(`id_marca`,`descripcion`,`id_usu`)VALUES('" & Format(rst2("id"), "00000") & "','" & rst2("name") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If
End Sub

Public Sub migrar_linea()

strCadena = "SELECT * FROM ginsac_grupo order by id"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`produccion`,`id_usu`)VALUES('" & Format(rst2("id"), "00000") & "','" & rst2("name") & "','no','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If
End Sub


Private Sub cmdactuallizar_insumos_Click()
strCadena = "select * from producto WHERE id_sublinea='00006' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "UPDATE producto SET insumo='si' WHERE id_linea='00012' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub

Private Sub persona_existe(ByVal in_dni As String, ByVal in_razon As String, ByVal in_direccion As String)
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(in_dni) & "'  LIMIT 1"
    Call ConfiguraRst(strCadena)
                      If rst.RecordCount < 1 Then
                        strCadena = "call P_insert_persona_ii('" & in_dni & "' " & _
                        ",'-', " & _
                        "'-' " & _
                        ",'-' " & _
                        ",'" & in_razon & "' " & _
                        ",'" & in_direccion & "' " & _
                        ",'-' " & _
                        ",'-'" & _
                        ",'no' " & _
                        ",'no'" & _
                        ",'no' " & _
                        ",'no' " & _
                        ",'no' " & _
                        ",'no' " & _
                        ",'si' " & _
                        ",'" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                        
                        
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & in_dni & "' LIMIT 1 "
                Call ConfiguraRstLocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen)VALUES ('" & in_dni & "','si','" & KEY_RUC & "','00001')"
                    CnBd.Execute (strCadena)
                End If
       End If
       
End Sub

Private Sub cmdcompras_Click()


strCadena = "select ai.id, ai.currency_id, ai.account_id, ai.state, ai.type, ai.date_invoice, ai.move_id, ai.supplier_number_part1,ai.supplier_number_part2, ai.nro_invoice, ai.amount_untaxed, ai.amount_tax, ai.amount_total, ai.residual,ai.move_name , ai.partner_id, ai.internal_number, aa.code, rp.Name, rp.ref from account_invoice ai inner join account_account aa on ai.account_id = aa.id inner join res_partner rp on ai.partner_id = rp.id where ai.company_id = 1 and ai.type = 'in_invoice' and ai.state = 'open' order by ai.partner_id, ai.date_invoice "
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
        
        If rst2("currency_id") = "3" Then ' dolares
           in_moneda = "00002"
        Else
           in_moneda = "00001"
        End If
    
            If Mid(Trim(rst2("internal_number")), 1, 1) = "B" Then
                     in_doc = "0003"
                     in_documento = ""
            End If
            If Mid(Trim(rst2("internal_number")), 1, 1) = "F" Then
                in_doc = "0001"
                in_documento = ""
            End If
            
            If Mid(Trim(rst2("internal_number")), 2, 1) = "N" Then
                in_doc = "0007"
                in_documento = ""
            End If
            
            
            If in_doc <> "00" Then
            
            
            numero = Split(Trim(rst2("internal_number")), "-")
            If Len(Trim(rst2("internal_number"))) < 10 Then
                'ELECTRONICO
                in_serie = Trim(numero(0))
                in_numero = Format(numero(1), "000000")
            Else
                'NO ELECTRONICO
                'in_serie = Mid(numero(0), 4, 3)
                If UBound(numero) > 1 Then
                    in_serie = numero(0) & "-" & numero(1)
                    in_numero = Format(numero(2), "000000")
                Else
                    in_serie = numero(0)
                    in_numero = Format(numero(1), "000000")
                End If
                
                
            End If
            in_documento = in_documento & in_serie & "-" & in_numero
            
            in_anulado = "no"
            forma_pago = "01"
            Select Case rst2("state")
                Case "paid"
                    forma_pago = "01"
                    in_saldo = 0
                Case "open"
                    forma_pago = "02"
                    in_saldo = rst2("amount_total")
                Case "cancel"
                    in_anulado = "si"
                    in_saldo = 0
            End Select
            
            id_cliente = get_id_cliente(rst2("partner_id"))
            If Len(Trim(id_cliente)) > 11 Then
                id_cliente = Mid(id_cliente, 1, 11)
            End If
            
            strCadena = "SELECT nombre_completo,direccion FROM persona where dni='" & id_cliente & "' LIMIT 1"
            Call ConfiguraRstLocal(strCadena)
            If rstLocal.RecordCount > 0 Then
                in_cliente = rstLocal("nombre_completo")
                in_direccion = rstLocal("direccion")
            Else
                in_cliente = "SIN DNI/RUC"
                in_direccion = "SIN REGISTRO"
               
            End If
        in_alm = "00001"
            
        If rst2("amount_tax") = 0 Then
           in_exonerado = rst2("amount_untaxed")
        Else
           in_exonerado = 0
        End If
        End If
        strCadena = "SELECT * FROM con_periodo WHERE mes='" & Month(rst2("date_invoice")) & "' and Ejercicio='" & Year(rst2("date_invoice")) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           in_periodo = rst("id")
        End If
        in_cta_compra = "0"
        KEY_usuario = "47541050"
        strCadena = "call P_insert_compra_ultimate('" & in_doc & "','" & in_alm & "','" & Format(rst2("date_invoice"), "YYYY-mm-dd") & "','" & Format(rst2("date_invoice"), "YYYY-mm-dd") & "','02'," & _
        "'01','--','" & in_moneda & "','" & Format(Month(rst2("date_invoice")), "00") & "','" & Year(rst2("date_invoice")) & "','" & Trim(in_serie) & "'," & _
        "'" & Format(Trim(in_numero), "00000000") & "','6','" & id_cliente & "','" & in_cliente & "','" & cambio_venta(rst2("date_invoice")) & "'," & _
        "'0','" & Val(rst2("amount_untaxed")) & "','" & Val(rst2("amount_tax")) & "','0','0','0','0','" & in_exonerado & "','0','" & Val(rst2("amount_total")) & "','" & Val(rst2("residual")) & "'," & _
        " '" & KEY_usuario & "','-','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_usuario & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
        strCadena = "UPDATE movimiento_compra SET migrado='si' WHERE id_compra='" & id_compra & "'"
        CnBd.Execute (strCadena)
        rst2.MoveNext
        
   Next i
End If
End Sub

Private Sub cmdCuentaCobrar_Click()
Dim in_fecha() As String
Dim in_vence() As String
Dim in_emision As String
'For i = 0 To 15000
'in_razon = Trim(Me.HfCuentasCobrar.TextMatrix(i, 9))
'in_direccion = Trim(Me.HfCuentasCobrar.TextMatrix(i, 10))
'in_dni = Trim(Me.HfCuentasCobrar.TextMatrix(i, 8))
 '      If Len(in_dni) < 8 Then
 '           in_dni = Format(in_dni, "00000000")
  '     End If
   '    Call persona_existe(in_dni, in_razon, in_direccion)
       'DoEvents
'Next i
        
For i = 0 To 15000
    If Trim(Me.HfCuentasCobrar.TextMatrix(i, 5)) <> "" Then
        in_dni = Trim(Me.HfCuentasCobrar.TextMatrix(i, 8))
        If Len(in_dni) < 8 Then
            in_dni = Format(in_dni, "00000000")
        End If
        in_emision = Format(Me.HfCuentasCobrar.TextMatrix(i, 0), "YYYY-mm-dd")
        in_vencimiento = Format(Me.HfCuentasCobrar.TextMatrix(i, 1), "YYYY-mm-dd")
        in_numero = Format(Me.HfCuentasCobrar.TextMatrix(i, 7), "000000")
        
        
        in_doc = Format(Trim(Me.HfCuentasCobrar.TextMatrix(i, 5)), "0000")
        in_serie = Me.HfCuentasCobrar.TextMatrix(i, 6)
        
        
        in_razon = Replace(Trim(Me.HfCuentasCobrar.TextMatrix(i, 9)), "'", " ")
        in_direccion = Trim(Me.HfCuentasCobrar.TextMatrix(i, 10))
        in_valor_venta = Val(Me.HfCuentasCobrar.TextMatrix(i, 11))
        in_exonerado = Val(Me.HfCuentasCobrar.TextMatrix(i, 11))
        in_igv = 0
        in_total = Val(Me.HfCuentasCobrar.TextMatrix(i, 13))
        in_saldo = Val(Me.HfCuentasCobrar.TextMatrix(i, 14))
        in_tc = 3.28
        
        in_cuota = Me.HfCuentasCobrar.TextMatrix(i, 2)
        If in_cuota <> "" Then
           Dim cuotas() As String
        End If
        
        cuotas = Split(in_cuota, "-")
        
        
        in_abreviatura = ""
           If in_doc = "0003" Then
              in_abreviatura = "BOLETA:"
           End If
           If in_doc = "0001" Then
              in_abreviatura = "FACTURA:"
           End If
           
           If in_doc = "0007" Then
              in_abreviatura = "NOTA CREDITO:"
           End If
           
           
       documento = in_abreviatura & in_serie & "-" & in_numero
       in_observacion = " -"
        If in_doc <> "0003" And in_doc <> "0001" And iin_doc <> "0007" Then
            x = 0
        End If
        Dim in_ventacloud As Double
        strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' and id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
           
            strCadena = "p_insert_venta_cabecera_ultimate_ii('" & in_doc & "','00001','02','00001','no'," & _
            "'" & in_serie & "','" & in_numero & "','" & in_dni & "','" & in_razon & "','" & in_valor_venta & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','" & in_saldo & "'," & _
            "'" & in_total & "','0','" & in_emision & "','" & in_vencimiento & "','00001','42546269','42546269','" & in_tc & "','" & dfac & "','" & Format(Month(in_emision), "00") & "','" & Year(in_emision) & "'" & _
            ",'" & documento & "','09:01 pm','T','" & in_direccion & "','no','-','-','0','-','0','-','00001','01','0','" & in_observacion & "','no','no','0','0','0','1','0','0','" & KEY_RUC & "')"
            Call ConfiguraRst(strCadena)
            id_venta1 = rst("in_venta")
            strCadena = "UPDATE movimiento_venta SET migrado='si' WHERE id_venta='" & id_venta1 & "'"
            CnBd.Execute (strCadena)
            
       
           strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & id_venta1 & "' LIMIT 1"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount < 1 Then
               strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,ruc)VALUES " & _
               "('" & id_venta1 & "','02','370','" & in_total & "','" & in_total & "','00','-','-','0','-','10111','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
           End If
          If in_cuota <> "" Then
             documento = "LETRA COBRAR:" & "999" & "-" & cuotas(2)
             in_observacion = " -"
             in_doc = "0412"
           
             strCadena = "p_insert_venta_cabecera_ultimate_ii('" & in_doc & "','00001','02','00001','no'," & _
             "'999','" & cuotas(2) & "','" & in_dni & "','" & in_razon & "','" & in_valor_venta & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','" & in_saldo & "'," & _
             "'" & in_total & "','0','" & in_emision & "','" & in_vencimiento & "','00001','42546269','42546269','" & in_tc & "','" & dfac & "','" & Format(Month(in_emision), "00") & "','" & Year(in_emision) & "'" & _
             ",'" & documento & "','09:01 pm','T','" & in_direccion & "','no','-','-','0','-','0','-','00001','01','0','" & in_observacion & "','no','no','0','0','0','1','0','0','" & KEY_RUC & "')"
             Call ConfiguraRst(strCadena)
             id_venta = rst("in_venta")
             strCadena = "UPDATE movimiento_venta SET migrado='si',id_referencia='" & id_venta1 & "' WHERE id_venta='" & id_venta & "'"
             CnBd.Execute (strCadena)
             strCadena = "UPDATE movimiento_venta SET migrado='si',saldo=0 WHERE id_venta='" & id_venta1 & "'"
             CnBd.Execute (strCadena)
             
          Else
              strCadena = "UPDATE movimiento_venta SET migrado='si',id_referencia='" & id_venta1 & "' WHERE id_venta='" & id_venta1 & "'"
              CnBd.Execute (strCadena)
          End If
        Else
           'strCadena = "DELETE FROM  movimiento_venta_cuotas where id_venta='" & rst("id_venta") & "' AND  ruc='20128836251'"
           'CnBd.Execute (strCadena)
           in_ventacloud = rst("id_venta")
           
           If in_cuota <> "" Then
             documento = "LETRA COBRAR:" & "999" & "-" & cuotas(2)
             in_observacion = " -"
             in_doc = "0412"
             
            
             strCadena = "p_insert_venta_cabecera_ultimate_ii('" & in_doc & "','00001','02','00001','no'," & _
             "'999','" & cuotas(2) & "','" & in_dni & "','" & in_razon & "','" & in_valor_venta & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','" & in_saldo & "'," & _
             "'" & in_total & "','0','" & in_emision & "','" & in_vencimiento & "','00001','42546269','42546269','" & in_tc & "','" & dfac & "','" & Format(Month(in_emision), "00") & "','" & Year(in_emision) & "'" & _
             ",'" & documento & "','09:01 pm','T','" & in_direccion & "','no','-','-','0','-','0','-','00001','01','0','" & in_observacion & "','no','no','0','0','0','1','0','0','" & KEY_RUC & "')"
             Call ConfiguraRst(strCadena)
             id_venta = rst("in_venta")
            
            strCadena = "UPDATE movimiento_venta SET migrado='si',id_referencia='" & in_ventacloud & "' WHERE id_venta='" & id_venta & "'"
             CnBd.Execute (strCadena)
            
            strCadena = "UPDATE movimiento_venta SET migrado='si',saldo=0 WHERE id_venta='" & in_ventacloud & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "SELECT sum(total) FROM movimiento_venta WHERE id_referencia='" & in_ventacloud & "'  and  ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            Dim in_acumulado As Double
            If IsNull(in_acumulado) = True Then
               in_acumulado = 0
            Else
               in_acumulado = rst(0)
            End If
             
            strCadena = "UPDATE movimiento_venta SET total='" & in_acumulado & "' WHERE id_venta='" & in_ventacloud & "'"
            CnBd.Execute (strCadena)
         
          End If
          
           
           
           
           
        End If
        End If
          
        '  DoEvents
siguiente:
      Next i
  
    
    
    
    




End Sub

Private Sub cmdKardex_Click()

End Sub

Private Sub cmdleerExcelCuentasCobrar_Click()
Dim Archivo As String
Archivo = Trim("Cuenta Cobrar" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.HfCuentasCobrar.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
      
      'Set obj = Nothing


End Sub

Private Sub cmdlinea_Click()


Dim in_seguro As String
Dim in_campo As String
strCadena = "DELETE FROM seguro_medico_detalle WHERE ruc='" & Trim(Me.txtruc.Text) & "'"
CnBd.Execute (strCadena)
'strCadena = "SELECT * FROM pa_aseguradoras WHERE nro_ruc!='" & Trim(Me.txtruc.Text) & "' "
'Call ConfiguraRstMigrar(strCadena)
If rstMigrar.RecordCount > 0 Then
    rstMigrar.MoveFirst
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rstMigrar.RecordCount - 1
    For i = 0 To rstMigrar.RecordCount - 1
         
         strCadena = "DELETE FROM persona WHERE dni='" & rstMigrar("nro_ruc") & "'"
         CnBd.Execute (strCadena)
         
         'If Len(Trim(rstMigrar("nro_ruc"))) > 7 Then
        ' If verificar_existencia("id_seguro", Format(rstMigrar("cod_aseg"), "00000"), "seguro_medico_detalle", Trim(Me.txtruc.Text)) = False Then
              strCadena = "INSERT INTO seguro_medico_detalle(`id_detalle`,`ruc_seguro`,`descripcion`,`detalle`,`fecha_registro`,`descuento_productos`,`factor`,`valor_consulta`,dominio,`eps`,`ruc`) " & _
            " VALUES('" & Format(rstMigrar("cod_aseg"), "00000") & "','" & Trim(rstMigrar("nro_ruc")) & "','" & Trim(rstMigrar("nom_aseg")) & "','" & rstMigrar("obser_aseg") & "',CURDATE(),'" & rstMigrar("des_productos") & "','" & rstMigrar("factor_aseg") & "','" & rstMigrar("valor_consulta") & "','-','" & rstMigrar("eps") & "','" & Trim(Me.txtruc.Text) & "')"
            CnBd.Execute (strCadena)
            If IsNull(rstMigrar("dir_aseg")) = False Then
                in_campo = Trim(rstMigrar("dir_aseg"))
            Else
                in_campo = "-"
            End If
           ' Call proceso_persona(Trim(rstMigrar("nro_ruc")), Trim(rstMigrar("nom_aseg")), "-", in_campo)
         'End If
        ' End If
         Me.ProgressBar1.Value = i
         DoEvents
         rstMigrar.MoveNext
    Next i
    
End If


End Sub

Private Sub cmdmigracionservicios_Click()
strCadena = "SELECT * FROM "
End Sub
Private Function get_dni(ByVal in_dni As String) As String

in_dni = Trim(in_dni)
in_numero = ""
For i = 0 To Len(in_dni) - 1
    in_numero2 = Mid(in_dni, Len(in_dni) - i, 1)
    If IsNumeric(in_numero2) = True Then
        If i = 0 Then
            in_numero = Mid(in_dni, Len(in_dni) - i, 1)
        Else
            in_numero1 = Mid(in_dni, Len(in_dni) - i, 1)
            in_numero = Trim(in_numero1 + in_numero)
        End If
    Else
        get_dni = in_numero
        Exit For
    End If
    
Next i

get_dni = in_numero
End Function
Private Sub cmdmigrar_clientes_Click()
Dim nombres() As String
Dim in_dni As String

GoTo empezar
strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "'  order by cod_unico ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rst("cod_unico") & "' "
        Call ConfiguraRstLocal(strCadena)
        If rstLocal.RecordCount < 2 Then
            strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rst("cod_unico") & "' and id_empresa='" & KEY_RUC & "' and id_personal='no' "
            Call ConfiguraRstLocal(strCadena)
            If rstLocal.RecordCount > 0 Then
                strCadena = "DELETE FROM entidad_empresa WHERE cod_unico='" & rst("cod_unico") & "' and id_empresa='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                strCadena = "DELETE FROM persona WHERE dni='" & rst("cod_unico") & "'"
                CnBd.Execute (strCadena)
            End If
        End If
        rst.MoveNext
        DoEvents
        Me.ProgressBar1.Value = i + 1
        Me.lblItem.Caption = Str(i)
   Next i
End If

empezar:


strCadena = "SELECT * FROM res_partner  order by id ASC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   Me.ProgressBar1.Min = 1
   Me.ProgressBar1.Max = rst2.RecordCount
   For i = 0 To rst2.RecordCount - 1
                
      
      If IsNull(rst2("ref")) = True Then
         GoTo siguiente
      Else
        in_dni = get_dni(Trim(rst2("ref")))
        If in_dni = "" Or in_dni = "00000000000" Then
            GoTo siguiente
        End If
      End If
      a_paterno = ""
      a_materno = ""
      nombre = ""
      If Len(in_dni) > 11 Then
        GoTo siguiente
      End If
      strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'  LIMIT 1"
      Call ConfiguraRst(strCadena)
      On Error GoTo siguiente
            If rst.RecordCount < 1 Then
                nombres = Split(rst2("name"), " ")
            On Error GoTo niii
                a_paterno = Trim(nombres(0))
                a_materno = Trim(nombres(1))
            
            If UBound(nombres()) > 3 Then
                nombre = nombres(2) & Space(1) & nombres(3)
            Else
                If UBound(nombres()) > 1 Then
                    nombre = nombres(2)
                End If
                If UBound(nombres()) >= 3 Then
                    nombre = nombres(2) & Space(1) & nombres(3)
                End If
                
            End If
niii:
            nombre_completo = Replace(Trim(rst2("name")), Chr(34), "")
            If IsNull(rst2("phone")) = True Then
                in_phone = "-"
            Else
                in_phone = rst2("phone")
            End If
            If IsNull(rst2("email")) = True Then
                in_mail = "-"
            Else
                in_mail = rst2("email")
            End If
            If IsNull(rst2("street")) = True Then
                in_direccion = "-"
            Else
                in_direccion = rst2("street")
            End If
            
                strCadena = "call P_insert_persona_ii('" & in_dni & "' " & _
                ",'" & Replace(UCase(a_paterno), "'", " ") & "', " & _
                "'" & Replace(UCase(a_materno), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(nombre)), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(nombre_completo)), "'", " ") & "' " & _
                ",'" & Replace(Trim(in_direccion), "'", "") & "' " & _
                ",'" & in_phone & "'" & _
                ",'" & in_mail & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
       Else
                
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & in_dni & "' LIMIT 1 "
                Call ConfiguraRstLocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen)VALUES ('" & in_dni & "','si','" & KEY_RUC & "','00001')"
                    CnBd.Execute (strCadena)
                Else
                    If IsNull(rst2("phone")) = False Then
                        strCadena = "UPDATE persona SET celular='" & in_phone & "' WHERE dni='" & in_dni & "'"
                        CnBd.Execute (strCadena)
                        strCadena = "INSERT INTO "
                    End If
                End If
             
                  
      End If

siguiente:
      
      
      rst2.MoveNext
      Me.ProgressBar1.Value = i + 1
      Me.lblItem.Caption = Str(i) & Space(10) & Str(rst2.RecordCount)
      DoEvents
   Next i
End If
'Call ventasss
End Sub
Private Sub pagos_ventas(ByVal in_referencia As String, ByVal in_alm As String, ByVal in_dni As String, ByVal n_cliente As String, ByVal in_venta_deuda As String, ByVal in_comprobante As String, ByVal in_moneda As String)
strCadena = "Select * from account_move_line  where ref='" & Trim(in_referencia) & "' and reconcile_partial_id>0 "
Call ConfiguraRstCloud(strCadena)
If rstCloud.RecordCount > 0 Then
   rstCloud.MoveFirst
   strCadena = "Select * from account_move_line  where reconcile_partial_id='" & Val(rstCloud("reconcile_partial_id")) & "' and ref<>'" & in_referencia & "' "
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 0 Then
      rstCloud.MoveFirst
      For i = 0 To rstCloud.RecordCount - 1
           in_observacion = "PAGO:" & in_comprobante
           Call generar_recibo(in_alm, in_moneda, in_dni, n_cliente, rstCloud("credit"), rstCloud("date_created"), in_observacion, in_venta_deuda)
           rstCloud.MoveNext
      Next i
   End If
End If

End Sub
Private Sub pagos_ventas_vargas(ByVal in_key As String)
strCadena = "Select * from qfacpag  where movkey='" & in_key & "' "
Call ConfiguraRstCloud(strCadena)
If rstCloud.RecordCount > 0 Then
   rstCloud.MoveFirst
   
      For i = 0 To rstCloud.RecordCount - 1
           in_observacion = "PAGO:" & in_comprobante
           Call generar_recibo(in_alm, in_moneda, in_dni, n_cliente, rstCloud("credit"), rstCloud("date_created"), in_observacion, in_venta_deuda)
           rstCloud.MoveNext
      Next i
   End If


End Sub

Private Sub generar_recibo(ByVal in_alm As String, ByVal in_moneda As String, ByVal in_dni As String, ByVal n_cliente As String, ByVal in_monto As Single, ByVal in_fecha As Date, ByVal in_observacion As String, ByVal in_venta_deuda As Double)

                    id_tipo_factura = "00001"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT * FROM  movimiento_venta WHERE serie='001' and  id_doc='0054'  ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstLocal(strCadena)
                    If rstLocal.RecordCount > 0 Then
                        in_numero = Format(Val(rstLocal("numero")) + 1, "000000")
                   
                        
                    End If
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    documento = Trim("RBO VENTA") & ":" & "001" & "-" & in_numero
                    strCadena = "P_insert_venta('0054','" & in_alm & "','01','" & in_moneda & "','no'," & _
                    "'001','" & in_numero & "','" & in_dni & "','" & n_cliente & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Format(in_fecha, "YYYY-mm-dd") & "','00001','42546269','42546269','3.28','" & dfac & "','" & Format(Month(in_fecha), "00") & "','" & Year(in_fecha) & "','" & documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','06439','" & Trim(in_observacion) & "','-','1','" & Val(in_monto) & "','0','" & Val(in_monto) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_venta SET id_comprobante='" & Val(in_venta_deuda) & "',observacion='" & Trim(in_observacion) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_venta SET saldo=saldo-'" & in_monto & "' WHERE id_venta='" & in_venta_deuda & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    
End Sub

Private Sub ventasss()
Dim numero() As String
strCadena = "SELECT * FROM account_invoice  where  company_id='3' and type='out_invoice' and state='open'   order by  date_invoice DESC, internal_number DESC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   
   
   For i = 0 To rst2.RecordCount - 1
            in_emision = Format(rst2("date_invoice"), "YYYY-mm-dd")
            in_doc = "00"
            
            
            
            If Mid(Trim(rst2("internal_number")), 1, 1) = "B" Then
                in_doc = "0003"
                in_documento = "BOLETA:"
            End If
            If Mid(Trim(rst2("internal_number")), 1, 1) = "F" Then
                in_doc = "0001"
                in_documento = "FACTURA:"
            End If
            
            If Mid(Trim(rst2("internal_number")), 2, 1) = "N" Then
                in_doc = "0007"
                in_documento = "N CREDITO:"
            End If
            
            
            If in_doc <> "00" Then
            
            
            numero = Split(Trim(rst2("internal_number")), "-")
            If Len(Trim(rst2("internal_number"))) < 10 Then
                'ELECTRONICO
                in_serie = Trim(numero(0))
                in_numero = Format(numero(1), "000000")
            Else
                'NO ELECTRONICO
                in_serie = Mid(numero(0), 4, 3)
                in_numero = Format(numero(1), "000000")
            End If
            in_documento = in_documento & in_serie & "-" & in_numero
            
            in_anulado = "no"
            forma_pago = "01"
            Select Case rst2("state")
                Case "paid"
                    forma_pago = "01"
                    in_saldo = 0
                Case "open"
                    forma_pago = "02"
                    in_saldo = rst2("amount_total")
                Case "cancel"
                    in_anulado = "si"
                    in_saldo = 0
            End Select
            
            in_alm = Format(rst2("shop_id"), "00000")
            
            If rst2("shop_id") = "3" Then
                in_alm = Format(1, "00000")
            End If
            If rst2("shop_id") = "4" Then
                in_alm = Format(2, "00000")
            End If
            
            
            If KEY_RUC <> "20493910052" Then
            If Val(in_alm) = 1 Then
               If forma_pago = "01" Then
                  id_forma_pago = "418"
               Else
                  id_forma_pago = "425"
               End If
               in_cuenta_caja = "10111"
            Else
                If forma_pago = "01" Then
                  id_forma_pago = "498"
               Else
                  id_forma_pago = "425"
               End If
               in_cuenta_caja = "10112"
            End If
            End If
            If KEY_RUC = "20493910052" Then
            If Val(in_alm) = 1 Then
               If forma_pago = "01" Then
                  id_forma_pago = "430"
               Else
                  id_forma_pago = "437"
               End If
               in_cuenta_caja = "1041"
            Else
                If forma_pago = "01" Then
                  id_forma_pago = "555"
               Else
                  id_forma_pago = "592"
               End If
               in_cuenta_caja = "10112"
            End If
            End If
            
            
            in_hash = rst2("digestvalue")

            id_cliente = get_id_cliente(rst2("partner_id"))
            If Len(Trim(id_cliente)) > 11 Then
                id_cliente = Mid(id_cliente, 1, 11)
            End If
            
            strCadena = "SELECT nombre_completo,direccion FROM persona where dni='" & id_cliente & "' LIMIT 1"
            Call ConfiguraRstLocal(strCadena)
            If rstLocal.RecordCount > 0 Then
                in_cliente = rstLocal("nombre_completo")
                in_direccion = rstLocal("direccion")
            Else
                in_cliente = "SIN DNI/RUC"
                in_direccion = "SIN REGISTRO"
               
            End If
            
            
            If rst2("company_id") = 1 Then
                in_total = rst2("amount_total")
                in_valor_venta = in_total / 1.18
                in_igv = in_total - in_valor_venta
                in_exonerado = 0
            Else
                in_total = rst2("amount_total")
                in_valor_venta = in_total
                in_igv = 0
                in_exonerado = in_valor_venta
            End If
            
            If rst2("currency_id") = 168 Then
               in_moneda = "00001"
            End If
            If rst2("currency_id") = 3 Then
               in_moneda = "00002"
            End If
            
            
            in_tipo_nota = ""
            in_motivo_nota = ""
            in_guia = "0"
            in_observacion = ""
            in_move = 0
            in_referencia = ""
            
            in_move = rst2("id")
            in_referencia = rst2("internal_number")
            
            
            

                   strCadena = "DELETE FROM movimiento_venta_monto_temporal WHERE ruc='" & KEY_RUC & "' and id_doc='" & in_doc & "' and serie='" & in_serie & "'and id_alm='" & in_alm & "' "
                   CnBd.Execute (strCadena)
                   
                   strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,cuotas,id_usuario,id_alm,fecha,cuenta_contable,ruc)VALUES " & _
                   "('" & in_doc & "','" & in_serie & "','" & in_numero & "','" & forma_pago & "','" & id_forma_pago & "','" & in_moneda & "','" & in_total & "','" & in_total & "','00','0','47541050','" & in_alm & "','" & Format(in_emision, "YYYY-mm-dd") & "','" & in_cuenta_caja & "','" & KEY_RUC & "')"
                   CnBd.Execute (strCadena)
                   
             
            
            
            
            
            
           'strCadena = "call CON_InsertaPeriodoNuevo('" & in_emision & "','" & KEY_RUC & "','42546269')"
           ' CnBd.Execute (strCadena)
            
            strCadena = "call put_ingreso_salida_migrar('42546269','00001','" & Format(in_emision, "YYYY-mm-dd") & "','01','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            strCadena = "p_insert_venta_migrar_ginsac('" & in_doc & "','" & in_alm & "','" & forma_pago & "','" & in_moneda & "','no'," & _
            "'" & in_serie & "','" & in_numero & "','" & id_cliente & "','" & in_cliente & "','" & in_valor_venta & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','" & in_saldo & "', " & _
            "'" & in_total & "','0','" & Format(in_emision, "YYYY-mm-dd") & "','" & Format(in_emision, "YYYY-mm-dd") & "','00001','47541050','47541050','3.28','no','" & Format(Month(in_emision), "00") & "','" & Year(in_emision) & "'" & _
            ",'" & in_documento & "','10:25','T','" & in_direccion & "','0','" & Trim(in_hash) & "','-','0','','0','-',' ','01', " & _
            "'0','-','no','si','1212','70111','0','0','0','" & in_documento & "','si','" & in_anulado & "','" & in_move & "','" & in_referencia & "','" & KEY_RUC & "')"
            Call ConfiguraRst(strCadena)
            id_venta = rst("in_venta")
            If IsNull(rst2("reference")) = False Then
                If rst2("state") = "open" Then
                Call pagos_ventas(rst2("reference"), in_alm, id_cliente, in_cliente, id_venta, in_documento, in_moneda)
                End If
            End If
                         
           
        
            
        
        
       
        
          

'StrNumero = Format(Trim(Str(Val(in_numero)) + 1), "000000")
'strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & in_doc & "' AND serie='" & in_serie & "'  AND ruc='" & Trim(KEY_RUC) & "'"
'CnBd.Execute (strCadena)
        
        
        
        
        End If
        
        
       rst2.MoveNext
       'DoEvents
   Next i
End If
MsgBox "ventas listas"

End Sub
Private Sub cmdtest_Click()

If Me.optmysql.Value = True Then
    Call conexion_cloud("01")
    Exit Sub
End If
If Me.optsql200.Value = True Then
    Call conexion_cloud("02")
End If

If Me.optsql2005.Value = True Then
    Call conexion_cloud("03")
End If

If Me.optpostgres.Value = True Then
    Call conexion_cloud("05")
End If

If Me.OptAccess.Value = True Then
    Call conexion_cloud("05")
End If

End Sub

Private Sub cmdunidades_Click()
Call migrar_producto_Vargas
'Call migrar_producto_Ginsac


MsgBox "MIGRACION EXITOSA", vbInformation

End Sub
Private Sub temporal(ByVal in_key As String, ByVal in_cliente As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String)
strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and dni_save='42546269'"
CnBd.Execute (strCadena)
strCadena = "select * from qfacdet a where a.`movkey`='" & in_key & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
For i = 0 To rst2.RecordCount - 1
    strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
    "('" & KEY_RUC & "','" & in_cliente & "','00001','" & in_doc & "','" & in_serie & "','" & in_numero & "','" & Format(rst2("movart"), "00000") & "','" & Val(rst2("movcan")) & "'," & _
    "'" & Val(rst2("movpre")) & " ','" & Val(rst2("movpre")) * Val(rst2("movcan")) & "','0','no','" & Trim(rst2("movnom")) & "','42546269')"
     CnBd.Execute (strCadena)
     rst2.MoveNext
Next i
End If
End Sub
Public Sub migrar_producto_Vargas()

Dim in_company As String
in_campany = 3

Dim in_cliente As String
Dim in_faltaba As String
GoTo iniciar2
'INICIAR MIGRACION DE COMPROBANTES VARGAS
iniciar:
strCadena = "SELECT `movkey`,`movtip`,`movser`,`movdoc`,`movanu`,`movfec`,`movven`,`movcli`,`movnom`,`movdom`,`movloc`,`movtel`,`movema`,`movdni`,`movruc`,`mov_oc`,`movper`,`movcdv`,`movcot`,`porigv`,`movmon`,`movtca`,`movbru`,`movdes`,`movsub`," & _
"`movint`,`movtot`,`movafe`,`movval`,`movimp`,`movpag`,`movdif`,`movtop`,`movsuc`,`movgui`,`mo2gui`,`porini`,`porint`,`movini`,`movmes`,`movcuo`,`movsal`,`movdir`,`movpla`,`movcue`,`movtdc`,`refkey`,`cheque` FROM qfacmov WHERE movkey='" & in_faltaba & "'"
GoTo siguiente2
iniciar2:
strCadena = "SELECT `movkey`,`movtip`,`movser`,`movdoc`,`movanu`,`movfec`,`movven`,`movcli`,`movnom`,`movdom`,`movloc`,`movtel`,`movema`,`movdni`,`movruc`,`mov_oc`,`movper`,`movcdv`,`movcot`,`porigv`,`movmon`,`movtca`,`movbru`,`movdes`,`movsub`," & _
"`movint`,`movtot`,`movafe`,`movval`,`movimp`,`movpag`,`movdif`,`movtop`,`movsuc`,`movgui`,`mo2gui`,`porini`,`porint`,`movini`,`movmes`,`movcuo`,`movsal`,`movdir`,`movpla`,`movcue`,`movtdc`,`refkey`,`cheque` FROM qfacmov WHERE movfec>='2018-01-01' and movfec<='2018-01-07' and movtip IN('07') ORDER BY movfec ASC"


siguiente2:
Call ConfiguraRstCloud(strCadena)
If rstCloud.RecordCount > 0 Then
   rstCloud.MoveFirst
   For i = 0 To rstCloud.RecordCount - 1
     
   
       in_serie = ""
        If rstCloud("movtip") = "01" Then
            in_serie = "F" & rstCloud("movser")
            in_documento = "FACTURA:" & rstCloud("movser") & "-" & rstCloud("movdoc")
            in_cliente = rstCloud("movruc")
            
        End If
        If rstCloud("movtip") = "03" Then
            in_serie = "B" & rstCloud("movser")
            in_documento = "BOLETA:" & in_serie & "-" & rstCloud("movdoc")
            in_cliente = rstCloud("movdni")
            If in_cliente = "" Then
                in_cliente = "00000000"
            End If
           
        End If
        
        
        
        If rstCloud("movtip") = "07" Then
            
            Dim in_referencia() As String
            in_referencia = Split(rstCloud("refkey"), "-")
            If in_referencia(0) = "01" Then
                      in_seriei = "F" & in_referencia(1)
             Else
                     in_seriei = "B" & in_referencia(1)
            End If
            strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & Format(in_referencia(0), "0000") & "' and serie='" & in_seriei & "' and numero='" & in_referencia(2) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                   Call ConfiguraRstLocal(strCadena)
                   If rstLocal.RecordCount < 1 Then
                      in_faltaba = rstCloud("refkey")
                      GoTo iniciar
                   
                   Else
                        in_cliente = rstLocal("id_cliente")
                        GoTo aaa
                   End If
            
             
            
            If rstCloud("movdni") <> "" Then
               in_cliente = rstCloud("movdni")
            Else
                in_cliente = rstCloud("movruc")
            End If
aaa:
            in_documento = "NOTA CREDITO:" & in_seriei & "-" & rstCloud("movdoc")
            
            
            
            
        End If
         Call temporal(rstCloud("movkey"), in_cliente, Format(rstCloud("movtip"), "0000"), in_serie, rstCloud("movdoc"))
         Call verificar_existencia_cliente_vargas(in_cliente, rstCloud("movnom"), rstCloud("movloc"))
            
            
                      
            
            
            strCadena = "call p_insert_venta_migrar_ginsac('" & Format(rstCloud("movtip"), "0000") & "','00001','01','00001','no'," & _
            "'" & in_serie & "','" & rstCloud("movdoc") & "','" & in_cliente & "','" & rstCloud("movnom") & "','" & rstCloud("movtot") & "','0','" & rstCloud("movtot") & "','" & rstCloud("movtot") & "','0', " & _
            "'" & rstCloud("movtot") & "','0','" & Format(rstCloud("movfec"), "YYYY-mm-dd") & "','" & Format(rstCloud("movven"), "YYYY-mm-dd") & "','00001','42546269','42546269','3.28','no','" & Format(Month(rstCloud("movfec")), "00") & "','" & Year(rstCloud("movfec")) & "'" & _
            ",'" & in_documento & "','10:25','T','" & rstCloud("movloc") & "','0','-','-','0','','0','-',' ','01', " & _
            "'0','-','no','si','1210101','7010101','0','0','0','" & in_documento & "','si','no','0','0','" & KEY_RUC & "')"
            Call ConfiguraRst(strCadena)
            id_venta = rst("in_venta")
            strCadena = "UPDATE movimiento_venta SET id_referencia='" & id_venta & "',migrado='si' WHERE id_venta='" & id_venta & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "SELECT * FROM qfacpag a where a.`movkey`='" & rstCloud("movkey") & "' LIMIT 1 "
            Call ConfiguraRst2(strCadena)
            If rst2.RecordCount > 0 Then
               rst2.MoveFirst
               For j = 0 To rst2.RecordCount - 1
                    in_tarjeta = "00"
                    If rst2("movnom") = "EFECTIVO SOLES" Then
                        in_forma_pago = "584"
                    End If
                    If rst2("movnom") = "VISA" Then
                        in_forma_pago = "608"
                        in_tarjeta = "01"
                    End If
                    If rst2("movnom") = "CHEQUE SOLES" Then
                        in_forma_pago = "591"
                    End If
                    If rst2("movnom") = "MASTER CARD" Then
                        in_forma_pago = "588"
                        in_tarjeta = "02"
                    End If
                    If rst2("movnom") = "DEPOSITOS" Then
                        in_forma_pago = "371"
                    End If
                    If rst2("movnom") = "NOTA DE CREDITO" Then
                        in_forma_pago = "372"
                    End If
                    If rst2("movnom") = "AMERCIAN EXPRESS" Then
                        in_forma_pago = "649"
                    End If
                    If rst2("movnom") = "BONIFICACION AL PERSONAL" Then
                        in_forma_pago = "584"
                    End If
                    If rst2("movnom") = "EFECTIVO DOLARES" Then
                        in_forma_pago = "607"
                    End If
                    If rst2("movnom") = "DINERS" Then
                        in_forma_pago = "649"
                    End If
              
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,cuenta_contable,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & in_forma_pago & "','" & rstCloud("movtot") & "','" & rstCloud("movtot") & "','00','1010101','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
               Next j
              End If
               If rstCloud("movtip") = "07" Then
                   
                   Dim in_ref() As String
                   in_ref = Split(rstCloud("refkey"), "-")
                   If in_ref(0) = "01" Then
                      in_seriei = "F" & in_ref(1)
                   Else
                     in_seriei = "B" & in_ref(1)
                   End If
                   
                   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & Format(in_ref(0), "0000") & "' and serie='" & in_seriei & "' and numero='" & in_ref(2) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                   Call ConfiguraRstLocal(strCadena)
                   If rstLocal.RecordCount > 0 Then
                      in_referencian = rstLocal("id_venta")
                   Else
                      in_referencian = 0
                   End If
                   strCadena = "UPDATE movimiento_venta SET id_comprobante='" & in_referencian & "',fecha_fact='" & Format(rstLocal("fecha_emision"), "YYYY-mm-dd") & "',id_doc_fact='" & rstLocal("id_doc") & "',serie_fact='" & rstLocal("serie") & "',numero_fact='" & rstLocal("numero") & "' WHERE id_venta='" & id_venta & "'"
                   CnBd.Execute (strCadena)
                   
                   
               End If
               
               
         
            
            
            
            
        rstCloud.MoveNext
                         
           
   Next i
End If


Exit Sub

GoTo producto
'strCadena = "DELETE FROM movimiento_venta where ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)

'strCadena = "DELETE from movimiento_venta_monto where ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)

strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM almacen_producto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

'strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)

'strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)

'strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)

'strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)

'strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "' "
'CnBd.Execute (strCadena)


strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM qmaeuni ORDER BY unicod"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
         strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & Format(i + 1, "00000") & "','" & UCase(rst2("uninom")) & "','" & UCase(rst2("unicod")) & "','" & KEY_RUC & "')"
         CnBd.Execute (strCadena)
         rst2.MoveNext
   Next i
End If




'Migracion de Clasificacion
 
                id_marca = "00000"
                strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','SIN MARCA','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
 
        


Dim id_linea As String
Dim id_sublinea As String

strCadena = "SELECT * FROM qmaelin ORDER BY lincod"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
        If Len(Trim(rst2("lincod"))) = 2 Then
           id_linea = Format(rst2("lincod"), "00000")
           strCadena = "SELECT * FROM linea WHERE id_linea = '" & id_linea & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount < 1 Then
               
               strCadena = "INSERT INTO linea(id_linea,descripcion,id_usu)VALUES('" & id_linea & "','" & rst2("linnom") & "','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
           End If
        End If
        rst2.MoveNext
   Next i
End If
   

strCadena = "SELECT * FROM qmaelin ORDER BY lincod"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
        If Len(Trim(rst2("lincod"))) > 2 Then
            id_linea = Format(Mid(rst2("lincod"), 1, 2), "00000")
            id_sublinea = Format(Mid(rst2("lincod"), 3, 2), "00000")
            strCadena = "SELECT * FROM linea_sub WHERE id_tipo = '" & id_sublinea & "' and id_linea='" & id_linea & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & rst2("linnom") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
        End If
        rst2.MoveNext
   Next i
        
 End If
        
        
            
producto:
       

strCadena = "SELECT `artcod`,`artnom`,`artlin`,`artmar`,`artuni`,`artfac`,`artun2`,`artmon`,`artpre`,`artpr2`,`artcos`,`artdes`,`artmt3`,`artara`,`artcue`,`artsta`,`artst1`,`artst2`,`artsto`,`artstb`,`artu01`,`artu02`,`artu03`,`artu04`,`artsol`,`artsoa`,`artsou`,`artdol`,`artdoa`,`artdou`,`artglo`,`ar3glo`,`artcds`,`artkit`,`artser`,`arttra`,`artlog`,`artigv`,`artisc`,`artptr`,artco1 FROM qmaeart WHERE artcod='14817' "
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
 rst2.MoveFirst
 For i = 0 To rst2.RecordCount - 1
        
      
        
        id_linea = Format(Mid(rst2("artlin"), 1, 2), "00000")
        in_sublinea = Format(Mid(rst2("artlin"), 3, 2), "00000")
        
        '----- MARCA *********
        in_marca = rst2("artmar")
        strCadena = "SELECT * FROM qmaemar WHERE marcod='" & in_marca & "'"
        Call ConfiguraRst3(strCadena)
        If rst3.RecordCount > 0 Then
            in_marca = rst3("marnom")
        End If
            
        strCadena = "SELECT * FROM marca WHERE descripcion = '" & in_marca & "'  AND id_usu='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM marca  where id_usu ='" & KEY_RUC & "' ORDER BY id_marca DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_marca = "00000"
                in_marca = "SIN MARCA"
            Else
                id_marca = Format(Val(rst("id_marca")) + 1, "00000")
            End If
            strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','" & in_marca & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_marca = rst("id_marca")
        End If
        
        'END  MARCA *********
        
        '----- COLOR *********
        id_color = "0000"
        in_tipo = "01"
        
        'END  MARCA *********
        '----- Unidad de medida *********
        in_unidad = Trim(rst2("artuni"))
            'in_unidad = rstCloud("write_uid")
            
        strCadena = "SELECT * FROM unidad WHERE abreviatura = '" & in_unidad & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            id_unidad = Format(rst("id_und"), "00000")
        Else
            id_unidad = Format(1, "00000")
        End If
            
       'END  MARCA *********
        nombre_producto = ""
        in_producto = UCase(Trim(rst2("artnom")))
        
        nombre_producto = Trim(in_producto)
        id_producto = Format(rst2("artcod"), "00000")
       
        If IsNull(rst2("artpre")) = True Then
            in_precio_venta = 0
        Else
            in_precio_venta = rst2("artpre")
        End If
        in_precio_venta = Format(in_precio_venta, "###0.00")
        in_precio_costo = Format(rst2("artco1"), "###0.00")
        in_precio_mayor = Format(rst2("artpr2"), "###0.00")
        
        strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
        "('" & id_producto & "','01','" & id_linea & "','" & in_sublinea & "','00001','" & id_color & "','" & nombre_producto & "','" & id_unidad & "','" & nombre_producto & "','" & id_marca & "','si','42546269','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For j = 0 To rst.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`) VALUES ('" & rst("id_alm") & "','" & id_producto & "','" & in_precio_venta & "','" & in_precio_costo & "','" & KEY_RUC & "','si')"
               CnBd.Execute (strCadena)
               rst.MoveNext
           Next j
        End If
        
siguiente:
      DoEvents
      
      rst2.MoveNext
    Next i
  End If
 

End Sub

Public Sub migrar_producto_Ginsac()

Dim in_company As String
in_campany = 3
GoTo mirarr

strCadena = "DELETE FROM movimiento_venta where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_venta_monto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM almacen_producto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)

strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)



strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM product_uom ORDER BY id"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
         strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & Format(rst2("id"), "00000") & "','" & UCase(rst2("name")) & "','" & UCase(rst2("name")) & "','" & KEY_RUC & "')"
         CnBd.Execute (strCadena)
         rst2.MoveNext
   Next i
End If
mirarr:

strCadena = "SELECT * FROM product_template t ,product_product p WHERE p.product_tmpl_id=t.id and t.company_id='" & in_campany & "' ORDER BY id ASC"
'strCadena = "SELECT * FROM product_product ORDER BY id DESC "
Call ConfiguraRstCloud(strCadena)
If rstCloud.RecordCount > 0 Then
 rstCloud.MoveFirst
'Me.ProgressBar1.Max = rst2.RecordCount
For i = 0 To rstCloud.RecordCount - 1
'reiiii:
       ' strCadena = "SELECT * FROM product_product WHERE id='8357' ORDER BY id DESC "
      '  Call ConfiguraRstCloud(strCadena)
        in_linea = Trim(rstCloud("group"))
               
        '----- CLASIFICACION *********
        If IsNull(in_linea) = True Then
            in_linea = "OTROS"
        End If
        
        strCadena = "SELECT * FROM linea WHERE descripcion = '" & in_linea & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_linea = "00001"
            Else
                id_linea = Format(Val(rst("id_linea")) + 1, "00000")
            End If
            
            
            strCadena = "INSERT INTO linea(id_linea,descripcion,id_usu)VALUES('" & id_linea & "','" & in_linea & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            id_sublinea = "00001"
            in_modelo = Trim(rstCloud("subgroup"))
            If IsNull(in_modelo) = True Then
                in_modelo = "SIN MODELO"
            End If
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        Else
            id_linea = rst("id_linea")
        End If
        id_sublinea = "00001"
        'END  CLASIFICACION *********
        
        
        
        
        If IsNull(in_modelo) = True Then
                in_modelo = "SIN MODELO"
        End If
        If in_modelo = "" Then
                in_modelo = "SIN MODELO"
        End If
         '----- MODELO *********
        strCadena = "SELECT * FROM linea_sub WHERE descripcion = '" & in_modelo & "' AND  id_linea='" & id_linea & "' AND id_usu='" & KEY_RUC & "' limit 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_sub where id_usu='" & KEY_RUC & "' ORDER BY id_tipo DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_sublinea = "00001"
            Else
                id_sublinea = Format(Val(rst("id_tipo")) + 1, "00000")
            End If
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
          id_sublinea = rst("id_tipo")
        End If
        'END  CLASIFICACION *********
        
        '----- MARCA *********
        id_marca = Format(rstCloud("marca_id"), "00000")
        If id_marca = "" Then
                id_marca = "00000"
        End If
            
        strCadena = "SELECT * FROM marca WHERE id_marca = '" & id_marca & "'  AND id_usu='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM marca  where id_usu ='" & KEY_RUC & "' ORDER BY id_marca DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_marca = "00000"
                in_marca = "SIN MARCA"
            Else
                id_marca = Format(Val(rst("id_marca")) + 1, "00000")
            End If
            strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','" & in_marca & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_marca = rst("id_marca")
        End If
        'END  MARCA *********
        
        '----- COLOR *********
        id_color = "0000"
        'in_color = Trim(Me.HfProductos.TextMatrix(i, 6))
        'strCadena = "SELECT * FROM imp_color WHERE descripcion LIKE '%" & in_color & "%'"
        'Call ConfiguraRst(strCadena)
        'If rst.RecordCount < 1 Then
        '    strCadena = "SELECT * FROM imp_color ORDER BY id_color DESC LIMIT 0,1"
        '    Call ConfiguraRst(strCadena)
        '    If rst.RecordCount < 1 Then
        '        id_color = "0000"
        '    Else
        '        id_color = Format(Val(rst("id_color")) + 1, "0000")
        '    End If
        '    strCadena = "INSERT INTO imp_color(id_color,descripcion)VALUES('" & id_color & "','" & in_color & "')"
        '    CnBd.Execute (strCadena)
        'Else
        '    id_color = rst("id_color")
        'End If
       
        in_tipo = rstCloud("tipo")
        If IsNull(in_tipo) = True Then
            in_tipo = "PRODUCTO"
        End If
        strCadena = "SELECT * FROM tipo_producto WHERE descripcion LIKE '%" & in_tipo & "%' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM tipo_producto where ruc='" & KEY_RUC & "' ORDER BY id_tipoproducto DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_tipo = "01"
            Else
                id_tipo = Format(Val(rst("id_tipoproducto")) + 1, "00")
            End If
            strCadena = "INSERT INTO tipo_producto(`id_tipoproducto`,`descripcion`,`ruc`)VALUES('" & id_tipo & "','" & in_tipo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_tipo = rst("id_tipoproducto")
        End If
        
        
        
        'END  MARCA *********
        '----- Unidad de medida *********
        id_unidad = Format(Trim(rstCloud("write_uid")), "00000")
        'in_unidad = rstCloud("write_uid")
        strCadena = "SELECT * FROM unidad WHERE id_und = '" & id_unidad & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            id_unidad = Format(1, "00000")
        End If
         '   strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 0,1"
          '  Call ConfiguraRst(strCadena)
          '  If rst.RecordCount > 0 Then
          '      id_unidad = Format(Val(rst("id_und") + 1), "00000")
          '  Else
          '     id_unidad = "00001"
          '  End If
          '  strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad & "','" & in_unidad & "','" & in_unidad & "','" & KEY_RUC & "')"
          '  CnBd.Execute (strCadena)
       ' Else
       '     id_unidad = rst("id_und")
       ' End If
        'END  MARCA *********
        nombre_producto = ""
        in_producto = UCase(Trim(rstCloud("name_template")))
        
        nombre_producto = Trim(in_producto)
       ' strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' ORDER BY id_producto DESC LIMIT 1"
       ' Call ConfiguraRst(strCadena)
       ' If rst.RecordCount > 0 Then
       '     id_producto = Format(Val(rst("id_producto") + 1), "00000")
       ' Else
       If IsNull(rstCloud("default_code")) = True Then
            GoTo siguiente:
       Else
            id_producto = Format(rstCloud("default_code"), "00000")
       End If
            
       ' End If
       ' If IsNull(rst2("list_price")) = True Then
            in_precio_venta = 0
       ' Else
        '    in_precio_venta = rst2("list_price")
        'End If
        in_precio_venta = Format(in_precio_venta, "###0.00")
        in_precio_costo = Format(0, "###0.00")
        in_precio_mayor = Format(0, "###0.00")
        On Error GoTo siguiente
        
        strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' and id_producto='" & id_producto & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 And Val(id_producto) > 0 Then
        'GoTo reiiii
        strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
        "('" & id_producto & "','" & id_tipo & "','" & id_linea & "','" & id_sublinea & "','00001','" & id_color & "','" & nombre_producto & "','" & id_unidad & "','" & nombre_producto & "','" & id_marca & "','si','42546269','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        
        strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For j = 0 To rst.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`) VALUES ('" & rst("id_alm") & "','" & id_producto & "','" & in_precio_venta & "','" & in_precio_costo & "','" & KEY_RUC & "','si')"
               CnBd.Execute (strCadena)
               rst.MoveNext
           Next j
        End If
        End If
siguiente:
      DoEvents
      
      rstCloud.MoveNext
    Next i
  End If
 

End Sub
Public Sub insertar_item(ByVal in_venta As String, ByVal in_cliente As String, ByVal in_alm As String, ByVal in_tipo_doc As String, ByVal in_serie As String, ByVal in_numero As String)
        Dim in_producto As String
        strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and id_doc='" & in_tipo_doc & "' and id_serie='" & in_serie & "' and dni_save='47541050' and id_alm='" & in_alm & "'"
        CnBd.Execute (strCadena)
        
        strCadena = "SELECT d.quantity,d.price_unit,d.price_subtotal,d.name,p.default_code FROM account_invoice_line d,product_product p WHERE d.product_id=p.id and  d.invoice_id='" & Val(in_venta) & "'"
        Call ConfiguraRstCloud(strCadena)
        If rstCloud.RecordCount > 0 Then
           rstCloud.MoveFirst
           For i = 0 To rstCloud.RecordCount - 1
                
                strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
                "('" & KEY_RUC & "','" & in_cliente & "','" & in_alm & "','" & in_tipo_doc & "','" & in_serie & "','" & in_numero & "','" & Format(rstCloud("default_code"), "00000") & "','" & Val(rstCloud("quantity")) & "'," & _
                "'" & Val(rstCloud("price_unit")) & " ','" & Val(rstCloud("price_unit")) * Val(rstCloud("quantity")) & "','0','si','" & Trim(rstCloud("name")) & "','47541050')"
                CnBd.Execute (strCadena)
            rstCloud.MoveNext
           Next i
        End If
        
        
      
    

End Sub
Public Sub insertar_item_venta(ByVal in_venta As String, ByVal in_cliente As String, ByVal in_alm As String, ByVal in_tipo_doc As String, ByVal in_serie As String, ByVal in_numero As String)


        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & in_venta & "' & " '"
        Call ConfiguraRstCloud(strCadena)
        If rstCloud.RecordCount > 0 Then
           rstCloud.MoveFirst
           For i = 0 To rstCloud.RecordCount - 1
                strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
                "('" & KEY_RUC & "','" & in_cliente & "','" & Format(in_alm, "00000") & "','" & in_tipo_doc & "','" & in_serie & "','" & in_numero & "','" & Format(rstCloud("product_id"), "00000") & "','" & Val(rstCloud("quantity")) & "'," & _
                "'" & Val(rstCloud("price_unit")) & " ','" & rstCloud("price_subtotal") & "','0','si','" & Trim(rstCloud("name")) & "','42546269')"
                CnBd.Execute (strCadena)
            rstCloud.MoveNext
           Next i
        End If
        
        
      
    

End Sub

Private Function get_id_cliente(ByVal in_partner) As String
strCadena = "SELECT * from  res_partner WHERE id='" & Val(in_partner) & "'"
Call ConfiguraRst3(strCadena)
If rst3.RecordCount > 0 Then
     If IsNull(rst3("ref")) = True Then
        get_id_cliente = "00000000"
     Else
        get_id_cliente = Trim(rst3("ref"))
     End If
        
           
    
Else
    get_id_cliente = "00000000"
End If
End Function
Private Function get_in_cliente(ByVal in_dni) As String
strCadena = "SELECT nombre_completo persona WHERE dni='" & Trim(in_dni) & "'"
Call ConfiguraRst(strCadena)
If rst3.RecordCount > 0 Then
    get_in_cliente = rst("nombre_completo")
Else
    get_in_cliente = "CLIENTE NO REGISTRADO"
End If
End Function
Private Sub cmdventas_Click()

Call ventasss



End Sub
Private Sub verificar_existencia_cliente(ByVal in_dni As String)

strCadena = "SELECT * FROM persona WHERE  dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRstLocal(strCadena)
If rstLocal.RecordCount > 0 Then
        strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(in_dni) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst2("a_paterno") & "', " & _
                "'" & rst2("a_materno") & "' " & _
                ",'" & rst2("nombres") & "' " & _
                ",'" & rst2("nombre_completo") & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & rst2("dni") & "' LIMIT 1 "
                Call ConfiguraRstLocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen)VALUES ('" & rst2("dni") & "','si','" & KEY_RUC & "','00001')"
                    CnBd.Execute (strCadena)
                End If
       End If
   
         '  DoEvents
   
   End If

End Sub
Private Sub verificar_existencia_cliente_vargas(ByVal in_dni As String, ByVal nombre As String, ByVal in_direccion As String)


        strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(in_dni) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & in_dni & "' " & _
                ",'-', " & _
                "'' " & _
                ",'' " & _
                ",'" & nombre & "' " & _
                ",'" & in_direccion & "' " & _
                ",'-' " & _
                ",'-'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & in_dni & "' LIMIT 1 "
                Call ConfiguraRstLocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen)VALUES ('" & in_dni & "','si','" & KEY_RUC & "','00001')"
                    CnBd.Execute (strCadena)
                End If
       End If


End Sub

Private Sub Command1_Click()
strCadena = "SELECT * FROM view_pacientes_v2 WHERE id_empresa='" & KEY_RUC & "' and id_cliente='si'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If Len(Trim(rst("dni"))) = 11 Then
       strCadena = "UPDATE entidad_empresa SET id_cliente='no',id_proveedor='si' WHERE cod_unico='" & rst("dni") & "' and id_empresa='" & Trim(Me.txtruc.Text) & "'"
       CnBd.Execute (strCadena)
       End If
       rst.MoveNext
   Next i
End If
End Sub

Private Sub Command10_Click()
strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and id_medico='si'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "DELETE FROM actividad_especialista WHERE dni='" & rst("cod_unico") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       strCadena = "DELETE FROM actividad_especialista_dia WHERE dni='" & rst("cod_unico") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub


Private Function get_seguro(ByVal in_seguro) As String
strCadena = "SELECT * FROM seguro_medico_detalle WHERE id_detalle='" & Format(in_seguro, "00000") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstLocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_seguro = rstLocal("descripcion")
End If
End Function
Private Function get_empleadora(ByVal in_empleadora As String) As String
strCadena = "SELECT * FROM pa_empleadoras WHERE cod_emp='" & in_empleadora & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   If IsNull(rst2("nro_ruc")) = True Then
        get_empleadora = "0"
   Else
        get_empleadora = rst2("nro_ruc")
   End If
   

End If
End Function

Private Sub Command12_Click()
    Dim in_unidad_frac As String
      
    strCadena = "SELECT * FROM fa_producto WHERE profrac='1'"
    Call ConfiguraRstMigrar(strCadena)
    If rstMigrar.RecordCount > 0 Then
       rstMigrar.MoveFirst
       For i = 0 To rstMigrar.RecordCount - 1
           in_unidad_frac = get_unidad(rstMigrar("codunidad_frac"))
           
           
            If rstMigrar("cod_minsa") = "0" Or IsNull(rstMigrar("cod_minsa")) = True Then
                in_minsa = Format(rstMigrar("codpro"), "000000")
            Else
                in_minsa = rstMigrar("cod_minsa")
           End If
           
           
           
           
           
           
           strCadena = "SELECT * FROM producto where id_producto='" & in_minsa & "' and ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst("nombre_prod") = rstMigrar("despro") Then
              
           
                strCadena = "UPDATE producto SET fraccionado='si' WHERE id_producto='" & in_minsa & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                strCadena = "UPDATE almacen_producto SET cantidad_fraccionada='" & rstMigrar("fracpro") & "',id_unidad_fraccion='" & in_unidad_frac & "' WHERE id_producto='" & in_minsa & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            Else
                MsgBox "Mira"
           End If
           rstMigrar.MoveNext
           DoEvents
       Next i
       MsgBox "listo"
    
    End If
    
    
    
           
           
End Sub

Private Sub Command14_Click()
Dim Archivo As String
Archivo = Trim("Producto Format" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.hfproductos.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
      
      'Set obj = Nothing

End Sub
Public Function get_cuenta_pago(ByVal in_registro As String) As String
On Error GoTo salir
strCadena = "SELECT id_cuenta_caja FROM forma_pago_detalle WHERE id_registro='" & in_registro & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstLocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_cuenta_pago = rstLocal("id_cuenta_caja")
Else
   get_cuenta_pago = 0
End If
Exit Function
salir:

End Function
Private Sub Command15_Click()

GoTo iniciar

strCadena = "SELECT * FROM movimiento_venta v, movimiento_venta_monto m WHERE  v.id_venta=m.id_venta and  v.id_doc IN( '0001','0003','0007') and v.ruc='" & KEY_RUC & "' ORDER BY v.id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
rst.MoveFirst
For i = 0 To rst.RecordCount - 1
 in_glosa = "PAGO :" & rst("documento")
               
              in_flujo = "1CIX000000000078"
              Call procesar_transaccion_venta(rst("id_alm"), get_cuenta_pago(rst("id_forma_pago")), Format(rst("fecha_emision"), "YYYY-mm-dd"), "00001", rst("id_cliente"), rst("ncliente"), in_glosa, rst("monto_caja"), "0", rst("id_venta"), "0", rst("documento"), Val(rst("tc")), Trim(rst("id_tarjeta_operacion")), "1CIX000000000174", in_flujo, rst("dni_save"), KEY_RUC)
              
          
            
      rst.MoveNext
Next i
End If
Exit Sub

KARDEX:
strCadena = "SELECT d.id_producto,v.id_venta,d.cantidad FROM movimiento_venta_detalle d,movimiento_venta v  WHERE v.id_doc IN('0001','0003','0007') and  d.id_venta=v.id_venta and  v.ruc='" & KEY_RUC & "' order by d.id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        Call actualizar_kardex(rst("id_producto"), rst("id_venta"), rst("cantidad"))
        rst.MoveNext
   Next i
End If
MsgBox "OKKKKKK KARDEX"

Exit Sub

iniciar:





For i = 0 To 10000
        
        If Me.hfproductos.TextMatrix(i, 0) >= 0 Then
        in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000")
        in_nombre_producto = Trim(Me.hfproductos.TextMatrix(i, 1))
        
        in_stock_cix = Val(Me.hfproductos.TextMatrix(i, 3))
        in_stock_piura = Val(Me.hfproductos.TextMatrix(i, 4))
        in_alm_cix = "00001" ' bagua
        in_alm_piura = "00002" ' nueva cajaramar
        GoTo actualizar_datos
        
        strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' order by id_producto DESC "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For j = 0 To rst.RecordCount - 1
               strCadena = "SELECT * FROM goreti_selva WHERE id_producto='" & rst("id_producto") & "'"
               Call ConfiguraRstLocal(strCadena)
               If rstLocal.RecordCount < 1 Then
                   strCadena = "DELETE FROM almacen_producto where id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                   CnBd.Execute (strCadena)
               End If
               rst.MoveNext
           Next j
        End If
        
        
        
       
        strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' and id_producto='" & in_producto & "'  LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
                                strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' and id_producto='" & in_producto & "'"
                                Call ConfiguraRst(strCadena)
                                If rst.RecordCount < 1 And Val(in_producto) > 0 Then
                                in_nombre_producto = Trim(Me.hfproductos.TextMatrix(i, 1))
                                strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
                                "('" & in_producto & "','00','00018','00175','00001','0000','" & in_nombre_producto & "','00001','" & in_nombre_producto & "','00000','si','42546269','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                                strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
                                Call ConfiguraRst(strCadena)
                                If rst.RecordCount > 0 Then
                                  rst.MoveFirst
                                   For j = 0 To rst.RecordCount - 1
                                       strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`) VALUES ('" & rst("id_alm") & "','" & in_producto & "','0','0','" & KEY_RUC & "','si')"
                                       CnBd.Execute (strCadena)
                                       rst.MoveNext
                                   Next j
                                End If
                                End If
                                GoTo actualizar_datos
      
       Else
                    strCadena = "UPDATE producto SET nombre_prod='" & in_nombre_producto & "' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
        End If
        GoTo siguiente

        in_precio_costo = Format(Me.hfproductos.TextMatrix(i, 8), "###0.00")
        in_precio_venta = Format(Me.hfproductos.TextMatrix(i, 9), "###0.00")
        in_precio_mayor = Format(Me.hfproductos.TextMatrix(i, 10), "###0.00")
        GoTo update_precios
        in_ubicacion_cix = Me.hfproductos.TextMatrix(i, 5)
        in_ubicacion_piura = Me.hfproductos.TextMatrix(i, 6)
        in_unidad = Trim(Me.hfproductos.TextMatrix(i, 2))
        
        
        
        
        in_stock_cix = Val(Me.hfproductos.TextMatrix(i, 3))
        in_stock_piura = Val(Me.hfproductos.TextMatrix(i, 4))
        
        strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' and descripcion='" & in_unidad & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            id_unidad = rst("id_und")
        Else
            id_unidad = ""
        End If
        
        If id_unidad <> "" And Len(id_unidad) > 3 Then
            strCadena = "UPDATE producto SET id_unidad='" & id_unidad & "' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        Else
                If id_unidad = "" Then
                   y = 0
                Else
                
                strCadena = "SELECT * FROM unidad WHERE  id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                   id_unidad = Format(Val(rst("id_und")) + 1, "00000")
                End If
                strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad & "','" & in_unidad & "','" & in_unidad & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "UPDATE producto SET id_unidad='" & id_unidad & "' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                End If
        End If
update_precios:
    strCadena = "UPDATE almacen_producto SET precio_compra='" & in_precio_costo & "',precio_venta='" & in_precio_venta & "',precio_mayor='" & in_precio_mayor & "' WHERE id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
        
    'strCadena = "UPDATE almacen_producto SET sector='" & in_ubicacion_cix & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm_piura & "'  and ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
    
    'strCadena = "UPDATE almacen_producto SET sector='" & in_ubicacion_piura & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm_cix & "'  and ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
actualizar_datos:
    strCadena = "DELETE FROM kardex WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "' "
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT d.id_producto,v.id_venta,d.cantidad FROM movimiento_venta_detalle d,movimiento_venta v  WHERE v.fecha_emision<='2017-12-31' and  id_producto='" & in_producto & "' and  v.id_doc IN('0001','0003','0007') and  d.id_venta=v.id_venta and  v.ruc='" & KEY_RUC & "' order by d.id_venta ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
        For k = 0 To rst.RecordCount - 1
            Call actualizar_kardex(rst("id_producto"), rst("id_venta"), rst("cantidad"))
            rst.MoveNext
        Next k
    End If
    
    strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & Val(in_precio_costo) & "','2018-01-02','" & in_alm_cix & "','" & Val(in_stock_piura) & "','42546269','PERCY RICARDO ANTICONA','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & Val(in_precio_costo) & "','2018-01-02','" & in_alm_piura & "','" & Val(in_stock_cix) & "','42546269','PERCY RICARDO ANTICONA','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT d.id_producto,v.id_venta,d.cantidad FROM movimiento_venta_detalle d,movimiento_venta v  WHERE v.fecha_emision>'2018-01-01' and  id_producto='" & in_producto & "' and  v.id_doc IN('0001','0003','0007') and  d.id_venta=v.id_venta and  v.ruc='" & KEY_RUC & "' order by d.id_venta ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
        For k = 0 To rst.RecordCount - 1
            Call actualizar_kardex(rst("id_producto"), rst("id_venta"), rst("cantidad"))
            rst.MoveNext
        Next k
    End If
  
        
        
        
    
    
    
      End If
siguiente:
      'DoEvents
      
      
    Next i
   'GoTo KARDEX:
    
 MsgBox "CORRECTO"

End Sub
Public Function procesar_transaccion_venta(ByVal in_alm As String, ByVal in_cta_origen As String, ByVal in_fecha As Date, ByVal in_tipo As String, ByVal in_proveedor As String, ByVal in_proveedor_des As String, ByVal in_observacion As String _
, ByVal in_monto As Single, ByVal in_cta_destino As String, ByVal in_venta As String, ByVal in_compra As String, ByVal in_documento As String, ByVal in_tc As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_tipo_flujo As String, ByVal in_dni_save As String, ByVal in_ruc As String) As Boolean
procesar_transaccion_venta = False


strCadena = "call sp_procesar_transaccion_caja_venta('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_tipo_flujo & "','" & in_ruc & "')"
CnBd.Execute (strCadena)


procesar_transaccion_venta = True

End Function
Private Sub actualizar_kardex(ByVal in_producto As String, ByVal in_venta As String, ByVal in_cantidad As Single)




End Sub


Private Sub Command16_Click()
Dim nombres() As String
Dim in_dni As String
strCadena = "SELECT * FROM Alumno order by dni ASC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   Me.ProgressBar1.Min = 1
   Me.ProgressBar1.Max = rst2.RecordCount
   For i = 0 To rst2.RecordCount - 1
    If IsNull(rst2("dni")) = True Then
         GoTo siguiente
      Else
        in_dni = Trim(rst2("dni"))
        If in_dni = "" Then
            GoTo siguiente
        End If
      End If
      a_paterno = ""
      a_materno = ""
      nombre = ""
      strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "' LIMIT 1"
      Call ConfiguraRst(strCadena)
      On Error GoTo siguiente
            If rst.RecordCount < 1 Then
                nombres = Split(rst2("name"), " ")
                a_paterno = Trim(nombres(0))
                a_materno = Trim(nombres(1))
            
            If UBound(nombres()) > 3 Then
                nombre = nombres(2) & Space(1) & nombres(3)
            Else
                If UBound(nombres()) > 1 Then
                    nombre = nombres(2)
                End If
                If UBound(nombres()) >= 3 Then
                    nombre = nombres(2) & Space(1) & nombres(3)
                End If
                
            End If
            
            nombre_completo = Replace(Trim(rst2("NomAlumno")), Chr(34), "")
            
                strCadena = "call P_insert_persona_ii('" & in_dni & "' " & _
                ",'" & Replace(UCase(a_paterno), "'", " ") & "', " & _
                "'" & Replace(UCase(a_materno), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(nombre)), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(nombre_completo)), "'", " ") & "' " & _
                ",'" & Replace(Trim(rst2("direccion")), "'", "") & "' " & _
                ",'" & rst2("TelAlumno") & "'" & _
                ",'-'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                strCadena = "UPDATE entidad_empresa SET codigo_universitario='" & rst2("CodigoU") & "' WHERE cod_unico='" & rst2("cod_unico") & "' and id_empresa='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
       Else
                If IsNull(rst2("phone")) = False Then
                strCadena = "UPDATE persona SET celular='" & rst2("TelAlumno") & "' WHERE dni='" & in_dni & "'"
                CnBd.Execute (strCadena)
                End If
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & in_dni & "' LIMIT 1 "
                Call ConfiguraRstLocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen,codigo_universitario)VALUES ('" & in_dni & "','si','" & KEY_RUC & "','00001','" & rst2("CodigoU") & "')"
                    CnBd.Execute (strCadena)
                End If
             
                  
      End If
      
siguiente:
      
      
      rst2.MoveNext
      Me.ProgressBar1.Value = i + 1
      Me.lblItem.Caption = Str(i) & Space(10) & Str(rst2.RecordCount)
      DoEvents
   Next i
End If

End Sub

Private Sub Command17_Click()

GoTo iniciar

strCadena = "SELECT * FROM movimiento_venta v, movimiento_venta_monto m WHERE  v.id_venta=m.id_venta and  v.id_doc IN( '0001','0003','0007') and v.ruc='" & KEY_RUC & "' ORDER BY v.id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
rst.MoveFirst
For i = 0 To rst.RecordCount - 1
 in_glosa = "PAGO :" & rst("documento")
               
              in_flujo = "1CIX000000000078"
              Call procesar_transaccion_venta(rst("id_alm"), get_cuenta_pago(rst("id_forma_pago")), Format(rst("fecha_emision"), "YYYY-mm-dd"), "00001", rst("id_cliente"), rst("ncliente"), in_glosa, rst("monto_caja"), "0", rst("id_venta"), "0", rst("documento"), Val(rst("tc")), Trim(rst("id_tarjeta_operacion")), "1CIX000000000174", in_flujo, rst("dni_save"), KEY_RUC)
              
          
            
      rst.MoveNext
Next i
End If
Exit Sub

KARDEX:
strCadena = "SELECT d.id_producto,v.id_venta,d.cantidad FROM movimiento_venta_detalle d,movimiento_venta v  WHERE v.id_doc IN('0001','0003','0007') and  d.id_venta=v.id_venta and  v.ruc='" & KEY_RUC & "' order by d.id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        Call actualizar_kardex(rst("id_producto"), rst("id_venta"), rst("cantidad"))
        rst.MoveNext
   Next i
End If
MsgBox "OKKKKKK KARDEX"

Exit Sub

iniciar:


in_limite_inf2 = 2
in_limite_sup2 = 99

'strCadena = "DELETE FROM almacen_producto_precio WHERE ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)

For i = 0 To 15000
        
        If Val(Me.hfproductos.TextMatrix(i, 0)) > 0 Then
        in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000")
        
        strCadena = "SELECT a.precio_compra FROM producto p,almacen_producto a WHERE  p.id_producto=a.id_producto and p.ruc=a.ruc and p.ruc='" & KEY_RUC & "' and p.id_producto='" & in_producto & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
                                strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' and id_producto='" & in_producto & "'"
                                Call ConfiguraRst(strCadena)
                                If rst.RecordCount < 1 And Val(in_producto) > 0 Then
                                in_nombre_producto = Trim(Me.hfproductos.TextMatrix(i, 1))
                                strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
                                "('" & in_producto & "','01','00004','00006','00001','0000','" & in_nombre_producto & "','00001','" & in_nombre_producto & "','00000','si','42546269','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                                strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`) VALUES ('00001','" & in_producto & "','0','0','" & KEY_RUC & "','si')"
                                CnBd.Execute (strCadena)
                               End If
                    
                            in_precio_costo = 0
        Else
            in_precio_costo = rst("precio_compra")
      
        End If
 
        

        

        in_precio_venta = Format(Me.hfproductos.TextMatrix(i, 3), "###0.00")
        
        
        in_venta2 = Format(Me.hfproductos.TextMatrix(i, 4), "###0.00")
        If in_venta2 > 0 Then
            strCadena = "INSERT INTO almacen_producto_precio(`id_alm`,`id_producto`,`precio`,`cant_ini`,`cant_fin`,`ruc`)VALUES('00001','" & in_producto & "','" & in_venta2 & "','" & in_limite_inf2 & "','" & in_limite_sup2 & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        
      
        
    in_stock_cix = Val(Me.hfproductos.TextMatrix(i, 5))
    strCadena = "UPDATE almacen_producto SET precio_venta='" & in_precio_venta & "',precio_mayor='" & in_venta2 & "' WHERE id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
        
  '  strCadena = "UPDATE almacen_producto SET sector='" & in_ubicacion_cix & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm_piura & "'  and ruc='" & KEY_RUC & "'"
  '  CnBd.Execute (strCadena)
    
  '  strCadena = "UPDATE almacen_producto SET sector='" & in_ubicacion_piura & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm_cix & "'  and ruc='" & KEY_RUC & "'"
  '  CnBd.Execute (strCadena)
        
    'in_producto
    
    strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & Val(in_precio_costo) & "','2018-01-01','00001','" & Val(in_stock_cix) & "','42546269','PERCY RICARDO ANTICONA','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
   
  
        
        
        
    
    
    
      End If
siguiente:
     ' DoEvents
      
      
    Next i
   ' GoTo KARDEX:
    
 MsgBox "CORRECTO"


End Sub

Private Sub Command18_Click()


strCadena = "SELECT d.id_producto,v.id_venta,d.cantidad FROM movimiento_venta_detalle d,movimiento_venta v  WHERE v.id_doc IN('0001','0003','0007') and  d.id_venta=v.id_venta and  v.ruc='" & KEY_RUC & "' order by d.id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        Call actualizar_kardex(rst("id_producto"), rst("id_venta"), rst("cantidad"))
        rst.MoveNext
   Next i
End If
End Sub

Private Sub Command19_Click()
Dim in_cantidad As Double
Dim in_costo As Double
Dim in_costo_ant As Double
Dim in_totalk As Double
Dim in_totalt As Double

Dim in_ini As Double
Dim in_ini_NEXT As Double
Dim in_ing As Double
Dim in_sal As Double
Dim in_saldo As Double
strCadena = "SELECT  DISTINCT id_producto from validar_kardex_v2    ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        'SALDO INICIAL
        
        strCadena = "SELECT IFNULL(SUM(cant_ingreso*costo_ingreso),0) FROM kardex_test_v3 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini1 = rstK(0)
        Else
            in_ini1 = 0
        End If
        
        strCadena = "SELECT IFNULL(sum(cant_ingreso*costo_ingreso),0) FROM kardex_test_v3 WHERE tipo IN ('02','18','21') and   id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini2 = rstK(0)
        Else
            in_ini2 = 0
        End If
        
        
        strCadena = "SELECT IFNULL(sum(cant_salida*costo_salida),0) FROM kardex_test_v3 WHERE tipo IN ('01','11','10') and   id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini3 = rstK(0)
        Else
            in_ini3 = 0
        End If
        
        in_ini = in_ini1 + in_ini2 + in_ini3
       
        'INGRESOS
        strCadena = "SELECT IFNULL(SUM(cant_ingreso*costo_ingreso),0) FROM kardex_test_v4 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        
        Call ConfiguraRstK(strCadena)
        in_ini_NEXT = rstK(0)
       
        
        
        If Int(in_ini) <> Int(in_ini_NEXT) Then
             MsgBox "PRODUCTO :" & rst("id_producto") + Chr(13) + Chr(13) + "INICIAL :" & Format(in_ini1, "#,##0.00") + Chr(13) + "INGRESOS:" & Format(in_ini2, "#,##0.00") + Chr(13) + "SALIDAS :" & Format(in_ini3, "#,##0.00") + Chr(13) + Chr(13) + Chr(13) & "SALDO FINAL :" & Space(5) & Format(in_ini, "#,##0.00") + Chr(13) + Chr(13) + "SALDO INICIO :" & Space(5) & Format(in_ini_NEXT, "#,##0.00")
             x = 0
        End If
        
        rst.MoveNext
        DoEvents
   Next i
End If

MsgBox "BUSQUEDA COMPLETA"

End Sub

Private Sub Command2_Click()
Dim in_codigo As String
GoTo nuevo
strCadena = "SELECT * FROM pa_tipos_serv where cod_tipo<>'' "
Call ConfiguraRstMigrar(strCadena)
If rstMigrar.RecordCount > 0 Then
   rstMigrar.MoveFirst
   For i = 0 To rstMigrar.RecordCount - 1
       in_tipo = "00"
       
       'For j = 1 To Len(Trim(rstMigrar("cod_tipo")))
        '   If Val(Mid(Trim(rstMigrar("cod_tipo")), j, 1)) = 0 Then
              in_linea = rstMigrar("cod_tipo")
         '     GoTo in_seg
         '  Else
         '     in_linea = Format(rstMigrar("cod_tipo"), "00000")
         '  End If
      ' Next j
'in_seg:
      strCadena = "SELECT * FROM linea WHERE id_linea='" & Trim(rstMigrar("cod_tipo")) & "' and id_usu='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount < 1 Then
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_tipo`,`nro_cuenta`,id_grupo_contable,`id_usu`) VALUES " & _
       "('" & rstMigrar("cod_tipo") & "','" & rstMigrar("nom_tipo") & "','" & in_tipo & "','" & Trim(rstMigrar("nro_cuenta_2011")) & "','" & rstMigrar("grupo_contable") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       End If
       rstMigrar.MoveNext
   Next i
End If


strCadena = "SELECT * FROM  producto WHERE id_tipo<>'07' and id_tipo<>'01' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "DELETE FROM producto WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
nuevo:
siguiente:
strCadena = "select * from pa_tarifa_servicios where cod_tipo='SOI'  ORDER BY cod_nivel ASC"
Call ConfiguraRstMigrar(strCadena)
If rstMigrar.RecordCount > 0 Then
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rstMigrar.RecordCount
   rstMigrar.MoveFirst
   For i = 0 To rstMigrar.RecordCount - 1
          in_codigo = rstMigrar("cod_nivel")
          strCadena = "SELECT * FROM producto where id_producto='" & in_codigo & "' and ruc='" & KEY_RUC & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount > 0 Then
            in_sub = get_sublineaA(rstMigrar("cod_grupo"), "00014")
            strCadena = "UPDATE producto SET id_sublinea='" & in_sub & "' WHERE id_producto='" & in_codigo & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
          Else
                If rstMigrar("estado") = "1" Then
                   in_habilitado = "si"
                Else
                   in_habilitado = "no"
                End If
                If IsNull(rstMigrar("nom_nuevo")) = True Then
                    comercial = "-"
                Else
                    comercial = Replace(rstMigrar("nom_nuevo"), "'", " ")
                End If
                
                strCadena = "INSERT INTO producto (id_producto,id_proveedor, id_unidad, id_linea,id_sublinea,nombre_prod,nombre_comercial,id_tipo,numero_placas,ruc) VALUES " & _
                " ('" & in_codigo & "','0','00001','" & rstMigrar("cod_tipo") & "','00000','" & Replace(rstMigrar("nom_tipo"), "'", " ") & "','" & comercial & "','02','" & rstMigrar("nro_placas") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                strCadena = "INSERT INTO almacen_producto (`id_alm`,`id_producto`,`precio_venta`,precio_convenio,`habilitado`,unidad,`ruc`)VALUES('00001','" & in_codigo & "','" & rstMigrar("particular") & "','" & rstMigrar("asegurado") & "','" & in_habilitado & "','" & rstMigrar("unidad") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
        End If
        
        rstMigrar.MoveNext
        Me.ProgressBar1.Value = i
        DoEvents
   Next i
   GoTo siguiente
End If
End Sub
Private Function get_sublineaA(ByVal in_grupo As String, ByVal in_linea As String) As String
strCadena = "SELECT * FROM pa_grupo_tarifario WHERE cod_grupo_tari='" & in_grupo & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    strCadena = "SELECT * FROM linea_sub WHERE id_tipo='" & Format(in_grupo, "00000") & "' and id_usu='" & KEY_RUC & "'"
    Call ConfiguraRstLocal(strCadena)
    If rstLocal.RecordCount < 1 Then
       strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES('" & Format(in_grupo, "00000") & "','" & in_linea & "','" & rst2("des_grupo_tari") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       get_sublineaA = Format(in_grupo, "00000")
    Else
       get_sublineaA = rstLocal("id_tipo")
    End If
End If
End Function



Private Sub Command20_Click()



For i = 0 To 15000
        
        If Val(Me.HfCuentasCobrar.TextMatrix(i, 0)) > 0 Then
          in_ruc = Trim(Me.HfCuentasCobrar.TextMatrix(i, 0))
          
        
          
          strCadena = "SELECT * from  persona WHERE dni='" & in_ruc & "' "
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
            x = 0
        Else
            in_cliente = rst("nombre_completo")
          
          End If
          
          in_comprobante = Trim(Me.HfCuentasCobrar.TextMatrix(i, 3))
          numero = Split(Trim(in_comprobante), "-")
          in_doc = Format(Trim(numero(0)), "0000")
          in_serie = Trim(numero(1))
          in_numero = Format(numero(2), "00000000")
          forma_pago = "02"
          
          If in_doc <> "0001" And in_doc <> "0003" And in_doc <> "0007" Then
            x = 0
          End If
         
          in_proveedor = in_ruc
          in_moneda = Format(Trim(Me.HfCuentasCobrar.TextMatrix(i, 4)), "00000")
          
          in_fecha_emision = Trim(Me.HfCuentasCobrar.TextMatrix(i, 1))
          in_fecha = Split(Trim(in_fecha_emision), "-")
          in_dia = Format(in_fecha(0), "00")
          in_mes = UCase(Format(Trim(in_fecha(1)), "00"))
          strCadena = "SELECT * FROM meses WHERE abreviatura='" & in_mes & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
            x = 0
          Else
            in_mes = rst("id_mes")
          End If
          
          in_anio = 2000 + in_fecha(2)
          fecha_emision = Format(Trim(in_anio & "-" & in_mes & "-" & in_dia), "YYYY-mm-dd")
         
         
         strCadena = "SELECT * FROM con_periodo WHERE mes='" & in_mes & "' and Ejercicio='" & in_anio & "'"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
           in_periodo = rst("id")
        Else
           x = 0
         End If
          
          in_fecha_vencimiento = Trim(Me.HfCuentasCobrar.TextMatrix(i, 2))
          in_fecha = Split(Trim(in_fecha_vencimiento), "-")
          in_dia = Format(in_fecha(0), "00")
          in_mes = Format(Trim(in_fecha(1)), "00")
          
          strCadena = "SELECT * FROM meses WHERE abreviatura='" & in_mes & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
            x = 0
          Else
            in_mes = rst("id_mes")
          End If
          
          in_anio = 2000 + in_fecha(2)
          fecha_vencimiento = Format(Trim(in_anio & "-" & in_mes & "-" & in_dia), "YYYY-mm-dd")
          
          
          
          
          
         
          in_alm = "00001"
          in_cta_compra = "0"
          in_usuario = "00122875"
          in_total = Abs(Val(Me.HfCuentasCobrar.TextMatrix(i, 5)))
          in_tc = Val(Me.HfCuentasCobrar.TextMatrix(i, 6))
          If in_tc = 0 Then
            in_tc = cambio_venta(fecha_emision)
          
          End If
            
       
          
          strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & Trim(Me.txtruc.Text) & "' and id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and id_proveedor='" & in_proveedor & "' LIMIT 1"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
             strCadena = "call P_insert_compra_ultimate('" & in_doc & "','" & in_alm & "','" & Format(fecha_emision, "YYYY-mm-dd") & "','" & Format(fecha_vencimiento, "YYYY-mm-dd") & "','02'," & _
             "'01','--','" & in_moneda & "','" & Format(Month(fecha_emision), "00") & "','" & Year(fecha_emision) & "','" & Trim(in_serie) & "'," & _
            "'" & Format(Trim(in_numero), "00000000") & "','6','" & in_ruc & "','" & in_cliente & "','" & in_tc & "'," & _
            "'0','0','0','0','0','0','0','0','0','" & Val(in_total) & "','0'," & _
            " '" & in_usuario & "','-','01','" & in_periodo & "','" & in_cta_compra & "','" & in_usuario & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
        
            strCadena = "UPDATE movimiento_compra SET migrado='si' WHERE id_compra='" & id_compra & "'"
            CnBd.Execute (strCadena)
            
            
          End If
         
            
            
            
        
        
        
   
        
        
        End If
 

Next i






End Sub

Private Sub Command21_Click()
Dim Archivo As String
Archivo = Trim("ctapagar" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.HfCuentasCobrar.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Hoja1")
      
      'Set obj = Nothing



End Sub

Private Sub Command22_Click()

strCadena = "SELECT DISTINCT id_producto FROM movimiento_venta_detalle WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT * FROM producto WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstLocal(strCadena)
       If rstLocal.RecordCount < 1 Then
          strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_producto='" & rst("id_producto") & "' and  ruc='" & KEY_RUC & "' ORDER BY  id_detalle_venta DESC LIMIT 1"
          Call ConfiguraRstP(strCadena)
          If rstP.RecordCount > 0 Then
          MsgBox "CODIGO:" & rst("id_producto") + Chr(13) + "producto:" + rstP("detalle") + Chr(13) + "Precio:" + Str(rstP("precio"))
          End If
       End If
       
       rst.MoveNext
   Next i
End If
End Sub

Private Sub Command23_Click()

strCadena = "SELECT * FROM "


End Sub

Private Function precio_inicial(ByVal in_producto As String) As Single

strCadena = "SELECT costo_promedio FROM kardex WHERE id_producto='" & in_producto & "' and  id_doc='0106' and ruc='" & Trim(Me.txtruc.Text) & "' ORDER BY fecha_emision ASC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   precio_inicial = rst(0)
End If

End Function


Private Sub Command25_Click()
strCadena = "SELECT * FROM producto WHERE   ruc='" & KEY_RUC & "' ORDER BY id_producto asc"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   For i = 0 To rst.RecordCount - 1
        Me.Command25.Caption = rst("id_producto")
     
            strCadena = "SELECT ifnull(stock_alm1,0) as stock_alm1 ,ifnull(stock_alm2,0) as stock_alm2 FROM ginsac WHERE id_producto='" & rst("id_producto") & "'  and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
               rstK.MoveFirst
               For j = 0 To rstK.RecordCount - 1
                   strCadena = "SELECT * FROM  kardex WHERE id_alm='00001' and id_doc='0106' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                   Call ConfiguraRstL(strCadena)
                   If rstL.RecordCount > 0 Then
                   If rstL.RecordCount = 1 And rstK("stock_alm1") = rstL("cantidad") Then
                   Else
                        MsgBox "SALDOS INCIALES DIFERENTES" + Chr(13) + "CANTIDAD SALDO:" & Str(rstK("stock_alm1")) + Chr(13) + "CANTIDAD KARDEX:" & rstL("cantidad")
                        x = 0
                   End If
                   End If
                   
                   
                   strCadena = "SELECT * from  kardex WHERE id_alm='00002' and id_doc='0106' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                   Call ConfiguraRstL(strCadena)
                   If rstL.RecordCount > 0 Then
                   If rstL.RecordCount = 1 And rstK("stock_alm2") = rstL("cantidad") Then
                   Else
                        MsgBox "SALDOS INCIALES DIFERENTES" + Chr(13) + "CANTIDAD SALDO:" & Str(rstK("stock_alm2")) + Chr(13) + "CANTIDAD KARDEX:" & rstL("cantidad")
                        x = 0
                   End If
                   End If
                   
                   rstK.MoveNext
                                      DoEvents
               Next j
            End If
       
        '****************************
        rst.MoveNext
        Me.Command25.Caption = rst("id_producto")
        DoEvents
        
        Next i
        End If
End Sub

Private Sub Command26_Click()


Call impresion_kardex_valorizado_demo(Me.DtpDesde.Value, Me.DtpHasta.Value, "00001")

End Sub
Public Sub imprecion_kardex_cabecera()
    
    Printer.Print Tab(0); "FORMATO 13.1 REGISTRO DEL INVENTARIO PERMANENTE VALORIZADO-DETALLE DEL INVENTARIO VALORIZADO"
    Printer.Print Tab(0); "PERIODO                                          :" & UCase(MonthName(Month(Me.DtpDesde))) & "-" & Year(Me.DtpDesde.Value)
    Printer.Print Tab(0); "RUC                                              :" & KEY_RUC
    Printer.Print Tab(0); "APELLIDOS Y NOMBRES, DENOMINACION O RAZON SOCIAL :" & Trim(Me.lblempresa.Caption)
    Printer.Print Tab(0); "ESTABLECIMIENTO                                  :" & Trim(Me.lbldireccion.Caption)
    Printer.Print Tab(0); "TIPO                                             :01-MERCADERIA"
    Printer.Print Tab(0); "METODO DE EVALUACION                             :PROMEDIO"
    Printer.Print Tab(0); "EXPRESDO EN                                      :SOLES"
    Printer.Print Tab(0); "==================================================================================================================================================="
    Printer.Print Tab(0); "                                                             ::::::    INGRESOS   ::::::     ::::::  SALIDAS   ::::::   ::::::     SALDOS    ::::::"
    Printer.Print Tab(0); "CODIGO   FECHA       TIPO    SERIE   NUMERO  T.OPERACION     CANT    COSTO.UNI   COSTO.T     CANT  COSTO.UNI COSTO.T    CANT    COSTO.UNI   COSTO.T"
    Printer.Print Tab(0); "==================================================================================================================================================="
End Sub
Public Sub impresion_kardex_valorizado_demo(ByVal fecha_ini As Date, ByVal fecha_fin As Date, ByVal in_alm As String)
    
    Dim sum_ing_cant As Double
    Dim sum_ing_tot As Double
    
    Dim sum_sali_cant As Double
    Dim sum_sali_tot As Double
    
    
    
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.CurrentX = 0
    Printer.CurrentY = 0
 
    Printer.ScaleWidth = 10#
    Printer.ScaleHeight = 28#
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    
    Printer.Font.Name = "Draft 17cpi"
    Printer.Font.Size = "10"
    
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
   Call imprecion_kardex_cabecera
    
    'strCadena = "SELECT DISTINCT movart,movnoa FROM qalmdet WHERE  movfec>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and movfec<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY movart ASC limit 40 "
    strCadena = "SELECT DISTINCT movart FROM qalmdet   ORDER BY movart ASC "
    If Month(fecha_ini) = 6 Then
        strCadena = "SELECT DISTINCT movart FROM qalmdet WHERE movart>='003169'   ORDER BY movart ASC "
    End If
    If Month(fecha_ini) = 10 Then
        strCadena = "SELECT DISTINCT movart FROM qalmdet WHERE movart>='009094'   ORDER BY movart ASC "
    End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
  
    For i = 0 To rst.RecordCount - 1
              If Val(Printer.CurrentY) >= 27.29045 Then
                   
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
              End If
                
        Printer.Print Tab(0); rst("movart") & Space(2) & "(7 UNIDADES)" & Space(2) & get_producto_vargas(rst("movart"))
        If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
        End If
                
        Printer.Print Tab(0); ""
        If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
        strCadena = "SELECT sum(if(movtip=2,movcan*-1,movcan)) FROM qalmdet WHERE movart='" & rst("movart") & "' and movfec<'" & Format(fecha_ini, "YYYY-mm-dd") & "' and movfec<='" & Format(fecha_fin, "YYYY-mm-dd") & "'"
        Call ConfiguraRstP(strCadena)
        in_cantidad_ini = rstP(0)
        
        strCadena = "SELECT movcos,movfec FROM qalmdet WHERE movart='" & rst("movart") & "' and movfec<'" & Format(fecha_ini, "YYYY-mm-dd") & "' and movfec<='" & Format(fecha_fin, "YYYY-mm-dd") & "' ORDER BY movfec ASC,movdo9 ASC LIMIT 1"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount > 0 Then
                in_costo_ini = rstP("movcos")
                ini_total_ini = in_cantidad_ini * in_costo_ini
                in_codigo = rst("movart")
                in_fecha = Format(fecha_ini, "dd-mm-YYYY")
                in_doc = Mid("" + Space(2), 1, 2)
                in_serie = Mid("" + Space(20), 4, 3)
                in_numero = Mid("" + Space(20), 8, 6)
                in_tipo = "16-SALDO INICIAL"
                
                ini_cantidad = Mid(Format(in_cantidad_ini, "#,##0.00") + Space(10), 1, 7)
                ini_costo = Mid(Format(in_costo_ini, "#,##0.00") + Space(10), 1, 6)
                ini_saldo = Mid(Format(in_cantidad_ini * in_costo_ini, "#,##0.00") + Space(10), 1, 10)
                
                ing_cantidad = Mid("" + Space(10), 1, 8)
                ing_costo = Mid(" " + Space(10), 1, 8)
                ing_total = Mid("" + Space(15), 1, 10)
                
                sal_cantidad = Mid(" " + Space(10), 1, 8)
                sal_costo = Mid(" " + Space(10), 1, 8)
                sal_total = Mid(" " + Space(15), 1, 10)
                
                sum_ing_cant = 0
                sum_sali_cant = 0
                
                sum_ing_tot = 0
                sum_sali_tot = 0
                
                Printer.Print Tab(0); rst("movart") & Space(2) & in_fecha & Space(3) & in_doc & Space(3) & in_serie & Space(2) & in_numero & Space(3) & Mid(in_tipo & Space(10), 1, 16) & ing_cantidad & ing_costo & ing_total & sal_cantidad & sal_costo & sal_total & sal_costo & Space(4) & Mid(Format(ini_cantidad, "#,##.00") + Space(15), 1, 10) & Mid(Format(ini_costo, "#,##0.00") + Space(15), 1, 10) & Mid(Format(ini_saldo, "#,##0.00") + Space(15), 1, 10)
                
        End If
        
        strCadena = "SELECT movart,movfec,movcom,movtip,movcan,movcos FROM qalmdet WHERE movart='" & rst("movart") & "' and movfec>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and movfec<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY movfec ASC,movdo9 ASC"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount > 0 Then
           rstP.MoveFirst
           For j = 0 To rstP.RecordCount - 1
                If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
                in_codigo = rst("movart")
                in_fecha = Format(rstP("movfec"), "dd-mm-YYYY")
                in_doc = Mid(rstP("movcom") + Space(2), 1, 2)
                in_serie = Mid(rstP("movcom") + Space(20), 4, 3)
                in_numero = Mid(rstP("movcom") + Space(20), 8, 6)
                If rstP("movtip") = "2" Then
                    in_tipo = "01-VENTA"
                Else
                    in_tipo = "02-COMPRA"
                End If
                'ingresos
                If rstP("movtip") = 1 Then
                              ing_cantidad = Mid(Format(rstP("movcan"), "#,##0.00") + Space(10), 1, 7)
                              ing_costo = Mid(Format(rstP("movcos"), "#,##0.00") + Space(10), 1, 6)
                              ing_total = Mid(Format(rstP("movcan") * rstP("movcos"), "#,##0.00") + Space(15), 1, 10)
                              sal_cantidad = Mid(" " + Space(10), 1, 8)
                              sal_costo = Mid(" " + Space(10), 1, 8)
                              sal_total = Mid(" " + Space(15), 1, 15)
                Else
                              ing_cantidad = Mid(" " + Space(10), 1, 7)
                              ing_costo = Mid(" " + Space(10), 1, 6)
                              ing_total = Mid(" " + Space(15), 1, 10)
                End If
                           'salidas
                If rstP("movtip") = 2 Then
                              sal_cantidad = Mid(Format(rstP("movcan"), "#,##0.00") + Space(10), 1, 8)
                              sal_costo = Mid(Format(rstP("movcos"), "#,##0.00") + Space(10), 1, 8)
                              sal_total = Mid(Format(rstP("movcan") * rstP("movcos"), "#,##0.00") + Space(15), 1, 15)
                              ing_cantidad = Mid(" " + Space(10), 1, 7)
                              ing_costo = Mid(" " + Space(10), 1, 6)
                              ing_total = Mid(" " + Space(15), 1, 10)
                Else
                              sal_cantidad = Mid(" " + Space(10), 1, 8)
                              sal_costo = Mid(" " + Space(10), 1, 8)
                              sal_total = Mid(" " + Space(15), 1, 15)
                End If
                           'saldos
                
                If j = 0 Then
                    
                    If rstP("movtip") = 2 Then
                        saldo_cantidad_saldo = Val(ini_cantidad) - rstP("movcan")
                        If Val(saldo_cantidad_saldo) = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = (Val(ini_saldo) + Val(sal_total)) / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    Else
                        saldo_cantidad_saldo = Val(ini_cantidad) + rstP("movcan")
                        If saldo_cantidad_saldo = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = (ini_saldo + Val(ing_total)) / saldo_cantidad_saldo
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    End If
                    
                    
                    
                
                Else
                    If rstP("movtip") = 2 Then
                        saldo_cantidad_saldo = Val(saldo_cantidad_saldo) - Val(rstP("movcan"))
                        If Val(Val(saldo_cantidad_saldo)) = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = saldo_total_saldo / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    Else
                        saldo_cantidad_saldo = Val(saldo_cantidad_saldo) + Val(rstP("movcan"))
                        If saldo_cantidad_saldo = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = saldo_total_saldo / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    End If
                    
                    
                    
                    
                End If
                
                Printer.Print Tab(0); rst("movart") & Space(2) & in_fecha & Space(4) & in_doc & Space(4) & in_serie & Space(4) & in_numero & Space(5) & Mid(in_tipo & Space(10), 1, 15) & Space(5) & ing_cantidad & ing_costo & ing_total & sal_cantidad & sal_costo & sal_total & Mid(Format(saldo_cantidad_saldo, "#,##.00") + Space(15), 1, 10) & Mid(Format(saldo_prom_saldo, "#,##0.00") + Space(15), 1, 10) & Mid(Format(saldo_total_saldo, "#,##0.00") + Space(15), 1, 10)
                
                sum_ing_cant = sum_ing_cant + Val(ing_cantidad)
                sum_sali_cant = sum_sali_cant + Val(sal_cantidad)
                
                sum_ing_tot = sum_ing_tot + Val(ing_total)
                sum_sali_tot = sum_sali_tot + Val(sal_total)
                
                sum_ing_cantn = Mid(Format(sum_ing_cant, "#,##0.00") + Space(10), 1, 8)
                sum_sali_cantn = Mid(Format(sum_sali_cant, "#,##0.00") + Space(15), 1, 10)
                              
                sum_ing_totn = Mid(Format(sum_ing_tot, "#,##0.00") + Space(10), 1, 8)
                sum_sali_totn = Mid(Format(sum_sali_tot, "#,##0.00") + Space(15), 1, 10)
                
                
                
                
                
                
                rstP.MoveNext
           Next j
           If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
           Printer.Print Tab(66); "-------" & Space(8) & "----------" & Space(8) & "-------" & Space(8) & "----------" & Space(10) & "----------"
           If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
           Printer.Print Tab(66); sum_ing_cantn & Space(6) & sum_ing_totn & sum_sali_cantn & Space(6) & sum_sali_totn
        End If
        
        
                           
        rst.MoveNext
    
    Next i
   End If
   
   
    
    
        
        Printer.EndDoc
        Exit Sub
 
End Sub


Private Sub Command27_Click()


Dim in_cantidad As Double
Dim in_costo As Double
Dim in_costo_ant As Double
Dim in_totalk As Double
Dim in_totalt As Double

Dim in_ini As Double
Dim in_ini_NEXT As Double
Dim in_ing As Double
Dim in_sal As Double
Dim in_saldo As Double
strCadena = "SELECT  DISTINCT id_producto from validar_kardex   ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        'SALDO INICIAL
        strCadena = "SELECT IFNULL(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v1 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        in_ini = rstK(0)
        in_ini_cant = rstK(1)
        'INGRESOS
        
        strCadena = "SELECT ifnull(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v1 WHERE tipo IN('02','18') and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        in_ing = rstK(0)
        in_ing_cant = rstK(1)
        'SALIDAS
        
        strCadena = "SELECT ifnull(sum(cantidad_salida*costo_salida),0),ifnull(sum(cantidad_salida),0) FROM kardex_test_v1 WHERE tipo='01' and id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        in_sal = rstK(0)
        in_sal_ant = rstK(1)
        
        in_saldo = in_ini + in_ing + in_sal
        
        
        
        strCadena = "SELECT ifnull(sum(cantidad_ingreso*costo_ingreso),0),ifnull(sum(cantidad_ingreso),0) FROM kardex_test_v2 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        in_ini_NEXT = rstK(0)
        in_ini_cant = rstK(1)
        
        
        If Int(in_saldo) <> Int(in_ini_NEXT) Then
             MsgBox "PRODUCTO :" & rst("id_producto") + Chr(13) + Chr(13) + "SALDO FINAL :" & Space(5) & Format(in_saldo, "#,##0.00") + Chr(13) + Chr(13) + "SALDO INICIO :" & Space(5) & Format(in_ini_NEXT, "#,##0.00")
        End If
        
        rst.MoveNext
        DoEvents
   Next i
End If

MsgBox "BUSQUEDA COMPLETA"



Exit Sub
strCadena = "call put_crear_kardex_producto_v3('" & KEY_RUC & "')"
CnBd.Execute (strCadena)
        
strCadena = "SELECT DISTINCT id_producto,id_alm FROM tmp_kardex_producto where  ruc='" & Trim(Me.txtruc.Text) & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & Trim(Me.txtruc.Text) & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstLocal(strCadena)
        If rstLocal.RecordCount > 0 Then
            rstLocal.MoveFirst
            For j = 0 To rstLocal.RecordCount - 1
                in_saldo = in_saldo + rstLocal("cantidad_real")
                If Val(in_saldo) <> rstLocal("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstLocal("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & Trim(Me.txtruc.Text) & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                End If
               
                rstLocal.MoveNext
                
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        rst.MoveNext
        DoEvents
        Me.Command27.Caption = Str(i) & Space(2) & rst.RecordCount
    Next i
End If
Exit Sub

End Sub

Private Sub Command28_Click()

Dim in_cantidad As Double
Dim in_costo As Double
Dim in_costo_ant As Double
Dim in_totalk As Double
Dim in_totalt As Double

Dim in_ini As Double
Dim in_ini_NEXT As Double
Dim in_ing As Double
Dim in_sal As Double
Dim in_saldo As Double
strCadena = "SELECT  DISTINCT id_producto from validar_kardex_v2    ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        'SALDO INICIAL
        
        strCadena = "SELECT IFNULL(saldo_stock*costo_saldo,0) FROM kardex_test_v3 WHERE id_alm='" & Trim(Me.txtAlmacen1.Text) & "' and   id_producto='" & rst("id_producto") & "' ORDER BY fecha DESC,id DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini1 = rstK(0)
        Else
            in_ini1 = 0
        End If
        
        strCadena = "SELECT IFNULL(saldo_stock*costo_saldo,0) FROM kardex_test_v3 WHERE id_alm='" & Trim(Me.txtAlmacen2.Text) & "' and   id_producto='" & rst("id_producto") & "' ORDER BY fecha DESC,id DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini2 = rstK(0)
        Else
            in_ini2 = 0
        End If
        
        
        strCadena = "SELECT IFNULL(saldo_stock*costo_saldo,0) FROM kardex_test_v3 WHERE id_alm='" & Trim(Me.txtAlmacen3.Text) & "' and   id_producto='" & rst("id_producto") & "' ORDER BY fecha DESC,id DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            in_ini3 = rstK(0)
        Else
            in_ini3 = 0
        End If
        
        in_ini = in_ini1 + in_ini2 + in_ini3
       
        'INGRESOS
        strCadena = "SELECT ifnull(sum(saldo_stock*costo_saldo),0) FROM kardex_test_v4 WHERE tipo='16' and  id_producto='" & rst("id_producto") & "'"
        Call ConfiguraRstK(strCadena)
        in_ini_NEXT = rstK(0)
       
        
        
        If Int(in_ini) <> Int(in_ini_NEXT) Then
             MsgBox "PRODUCTO :" & rst("id_producto") + Chr(13) + Chr(13) + "ALM 01 :" & Format(in_ini1, "#,##0.00") + Chr(13) + "ALM 02 :" & Format(in_ini2, "#,##0.00") + Chr(13) + "ALM 03 :" & Format(in_ini3, "#,##0.00") + Chr(13) + Chr(13) + Chr(13) & "SALDO FINAL :" & Space(5) & Format(in_ini, "#,##0.00") + Chr(13) + Chr(13) + "SALDO INICIO :" & Space(5) & Format(in_ini_NEXT, "#,##0.00")
             x = 0
        End If
        
        rst.MoveNext
        DoEvents
   Next i
End If

MsgBox "BUSQUEDA COMPLETA"

End Sub

Private Sub Command29_Click()
Dim in_id As String
strCadena = "SELECT * FROM con_cuentacontable WHERE Ejercicio='2017' ORDER BY ID ASC"
Call ConfiguraRstLocal(strCadena)
If rstLocal.RecordCount > 0 Then
    rstLocal.MoveFirst
    in_activo = rstLocal("Activo")
    in_tesoreria = rstLocal("Tesoreria")
    ID = 4412
    For i = 0 To rstLocal.RecordCount - 1
        ID = ID + 1
        in_id = "1CIX" & Format(Val(ID) + 1, "000000000000")
        strCadena = "INSERT INTO con_cuentacontable(id,`IdEmpresaSis`,`IdSucursal`,`IdNaturaleza`,`Ejercicio`,`NroCuenta`,`Descripcion`," & _
        "`MonedaExtranjera`,`IndCuentaDependiente`,`IdCuentaContableDepende`,`CtaCtbleDepende`,`IndMovimiento`,`DigitoSubfijo`,`CuentaSUNAT`, " & _
        " `IndFlujoCaja`,`IndConciliacion`,`IndDocumento`,`IndObligacion`,`IndDebe`,`IndHaber`,`IndGastoFuncion`,`IndItemGasto`,`IndCentroCosto`," & _
        " `IndTrabajador`,`IndBanco`,`Analisis01`,`Analisis02`,`Tesoreria`,`Activo`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`)VALUES  " & _
        "('" & in_id & "','" & KEY_RUC & "','00001','" & rstLocal("IdNaturaleza") & "','2018','" & rstLocal("NroCuenta") & "','" & rstLocal("Descripcion") & "' " & _
        ",'" & rstLocal("MonedaExtranjera") & "','" & rstLocal("IndCuentaDependiente") & "','" & rstLocal("IdCuentaContableDepende") & "', " & _
        "'" & rstLocal("CtaCtbleDepende") & "','" & rstLocal("IndMovimiento") & "','" & rstLocal("DigitoSubfijo") & "','" & rstLocal("CuentaSUNAT") & "' " & _
        ",'" & rstLocal("IndFlujoCaja") & "','" & rstLocal("IndConciliacion") & "','" & rstLocal("IndDocumento") & "','" & rstLocal("IndObligacion") & "','" & rstLocal("IndDebe") & "', " & _
        "'" & rstLocal("IndHaber") & "','" & rstLocal("IndGastoFuncion") & "','" & rstLocal("IndItemGasto") & "','" & rstLocal("IndCentroCosto") & "' " & _
        ",'" & rstLocal("IndTrabajador") & "','" & rstLocal("IndBanco") & "','" & rstLocal("Analisis01") & "','" & rstLocal("Analisis02") & "' " & _
        ",'" & rstLocal("Tesoreria") & "','" & rstLocal("Activo") & "','42546269',CURDATE(),'42546269',CURDATE())"
        CnBd.Execute (strCadena)
        rstLocal.MoveNext
    Next i
End If

End Sub

Private Sub Command3_Click()
strCadena = "SELECT DISTINCT cod_grupo FROM pa_tarifa_servicios"
Call ConfiguraRst(strCadena)

End Sub

Private Sub Command4_Click()
strCadena = "SELECT dni FROM persona WHERE dni<>'0' and dni<>'00000000' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rst("dni") & "' LIMIT 1"
        Call ConfiguraRstLocal(strCadena)
        If rstLocal.RecordCount < 1 Then
                If Len(rst("dni")) > 8 Then
                strCadena = "DELETE FROM persona WHERE dni='" & rst("dni") & "'"
                CnBd.Execute (strCadena)
                End If
        End If
        
        
        
        rst.MoveNext
        Me.ProgressBar1.Value = i
        DoEvents
    Next i
End If

End Sub

Private Sub Command5_Click()
strCadena = "select id_detalle from agenda WHERE ruc='20487911586'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
        strCadena = "DELETE FROM agenda WHERE id_detalle='" & rst("id_detalle") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Me.ProgressBar1.Value = i
        DoEvents
   Next i
End If

End Sub

Private Sub Command6_Click()
strCadena = "SELECT * FROM ambulancia_precios WHERE ruc='20487911586'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO ambulancia_precios(`descripcion`,`kilometros`,`precio`,`ruc`)VALUES('" & rst("descripcion") & "','" & rst("kilometros") & "','" & rst("precio") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub

Private Sub Command7_Click()
strCadena = "SELECT * FROM ron_hospitalizacion ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
        strCadena = "DELETE FROM ron_hospitalizacion WHERE id='" & rst("id") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Me.ProgressBar1.Value = i
        DoEvents
   Next i


End If
End Sub

Private Function get_unidad(ByVal in_unidad As String) As String

    
    strCadena = "SELECT * FROM Unidad WHERE CODIUNIDAD ='" & in_unidad & "'"
    Call ConfiguraRst2(strCadena)
    If rst2.RecordCount > 0 Then
        strCadena = "SELECT * FROM unidad WHERE descripcion='" & rst2("UNIDAD") & "' and id_usu='" & KEY_RUC & "' LIMIT 1 "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            get_unidad = rst("id_und")
        Else
            strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                in_codigo = Format(Val(rst("id_und")) + 1, "00000")
            Else
                in_codigo = Format(1, "00000")
            End If
    
            strCadena = "INSERT INTO unidad(`id_und`,`abreviatura`,`descripcion`,`id_usu`)VALUES('" & in_codigo & "','" & rst2("UNIDAD") & "','" & rst2("UNIDAD") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            get_unidad = in_codigo
        End If
        
        
    End If

End Function
Private Sub Command9_Click()
Dim in_controlado As String

'Call get_tipo_familia

strCadena = "SELECT id FROM producto WHERE id_tipo='07' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
    strCadena = "DELETE FROM producto WHERE id='" & rst("id") & "'"
    CnBd.Execute (strCadena)
    rst.MoveNext
   ' If i Mod 100 = 0 Then
   Me.lblcantidad.Caption = i
        DoEvents
   ' End If
    Me.ProgressBar1.Value = i
   Next i
End If



strCadena = "SELECT * FROM fa_producto"
Call ConfiguraRstMigrar(strCadena)
If rstMigrar.RecordCount > 0 Then
    'stMigrar.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rstMigrar.RecordCount
   
   For i = 0 To rstMigrar.RecordCount - 1
       If rstMigrar("controlado") = 1 Then
          in_controlado = "si"
       Else
          in_controlado = "no"
       End If
       If rstMigrar("estado") = "S" Then
          in_habilitado = "si"
       Else
          in_habilitado = "no"
       End If
       
       If rstMigrar("profrac") = 1 Then
           in_fraccionado = "si"
           in_unidad_frac = get_unidad(rstMigrar("codunidad_frac"))
           in_cantidad_fraccionada = rstMigrar("fracpro")
       Else
           in_fraccionado = "no"
           in_cantidad_fraccionada = 0
           in_unidad_frac = 0
       End If
   
   
           If rstMigrar("cod_minsa") = "0" Or IsNull(rstMigrar("cod_minsa")) = True Then
                in_minsa = 0
                in_interno = rstMigrar("codpro")
                in_digemid = "no"
           Else
                in_minsa = rstMigrar("cod_minsa")
                in_interno = rstMigrar("codpro")
                in_digemid = "si"
           End If
           
           If rstMigrar("codtip") = "I" Then
              in_insumo = "si"
           Else
              in_insumo = "no"
           End If
          
           in_costo_compra = rstMigrar("cost_comp_part")
           If rstMigrar("igvpro") > 0 Then
              in_igv = "si"
              in_valor_neto_part = Round(rstMigrar("cost_comp_part") * (1 + rstMigrar("utilidad_part") / 100), 2)
              in_precio_venta_part = in_valor_neto_part * 1.18
              in_utilidad_particular = rstMigrar("utilidad_part")
              
              in_utilidad_convenio = 33
              in_valor_neto_conv = Round(rstMigrar("costocom_parti") * (1.33), 2)
              in_precio_convenio = Round(rstMigrar("costocom_parti") * 1.33, 2) * 1.18
           Else
               in_utilidad_particular = rstMigrar("utilidad_part")
               in_utilidad_convenio = 33
               in_valor_neto_part = Round(rstMigrar("cost_comp_part") * (1 + rstMigrar("utilidad_part") / 100), 2)
               in_precio_venta_part = in_valor_neto_part
               
             in_valor_neto_conv = rstMigrar("costocom_parti") * 1.33
              in_precio_convenio = in_valor_neto_conv
              in_igv = "no"
           End If
           
           
           
           'strCadena = "SELECT * FROM producto WHERE id_producto='" & in_interno & "' and ruc='" & KEY_RUC & "'"
           'Call ConfiguraRst(strCadena)
           'If rst.RecordCount > 0 Then
            '    MsgBox rst("nombre_prod")
             '   strCadena = "DELETE FROM producto WHERE id_producto='" & in_interno & "' and ruc='" & KEY_RUC & "'"
              '  CnBd.Execute (strCadena)
           'End If
           strCadena = "INSERT INTO producto (id_producto,id_digemid,id_proveedor, id_unidad,id_forma_farm, id_linea,id_sublinea,nombre_prod,id_tipo,insumo,id_familia,fraccionado,id_interno,registro_digemid,producto_controlado,ruc) VALUES " & _
           " ('" & in_interno & "','" & in_minsa & "','" & get_proveedor(rstMigrar("codlab")) & "','" & get_unidad(rstMigrar("codiunidad")) & "','00000','00012','" & get_sublinea(Trim(rstMigrar("codtip"))) & "','" & Replace(rstMigrar("despro"), "'", " ") & "','07','" & in_insumo & "','" & Format(rstMigrar("codfam"), "00000") & "','" & in_fraccionado & "','" & in_interno & "','" & in_digemid & "','" & in_controlado & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "INSERT INTO almacen_producto (`id_alm`,`id_producto`,`valor_neto_part`,valor_neto_conv,`precio_venta_part`,precio_convenio,precio_compra,utilidad_particular,utilidad_convenio,cantidad_fraccionada,id_unidad_fraccion,`habilitado`,`ruc`) " & _
           "VALUES('00153','" & in_interno & "','" & in_valor_neto_part & "','" & in_valor_neto_conv & "','" & in_precio_venta_part & "','" & in_precio_convenio & "','" & in_costo_compra & "', " & _
           "'" & in_utilidad_particular & "','" & in_utilidad_convenio & "','" & in_cantidad_fraccionada & "','" & in_unidad_frac & "','" & in_habilitado & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
          
   '   End If
siguiente:
            Me.ProgressBar1.Value = i
            rstMigrar.MoveNext
            DoEvents
            Me.lblcantidad.Caption = i
       
   Next i
End If

End Sub

Public Function get_proveedor(ByVal deslab As String) As String
strCadena = "SELECT * FROM fa_laboratorios WHERE codlab='" & deslab & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    strCadena = "SELECT e.cod_unico FROM entidad_empresa e, persona p where e.cod_unico=p.dni and p.nombre_comercial LIKE '%" & rst2("deslab") & "%' and e.id_empresa='20103269319' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        get_proveedor = rst("cod_unico")
    Else
        strCadena = "SELECT e.cod_unico FROM entidad_empresa e, persona p where e.cod_unico=p.dni and p.nombre_completo LIKE '%" & rst2("deslab") & "%' and e.id_empresa='20103269319' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           get_proveedor = rst("cod_unico")
        Else
            get_proveedor = "0"
        End If
    End If
End If
End Function
Private Function get_producto() As String
strCadena = "SELECT id_producto FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY id DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   get_producto = rst("id_producto")
End If


End Function




Private Sub get_tipo_familia()
    strCadena = "DELETE FROM producto_familia WHERE ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)

       strCadena = "SELECT * FROM fa_familias "
       Call ConfiguraRst2(strCadena)
       If rst2.RecordCount > 0 Then
        For i = 0 To rst2.RecordCount - 1
        
            strCadena = "INSERT INTO producto_familia(id_familia,descripcion,ruc)VALUES('" & Format(Val(rst2("codfam")), "00000") & "','" & UCase(rst2("desfam")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            rst2.MoveNext
        Next i
       
    End If

End Sub

Private Function get_sublinea(ByVal in_codt As String) As String
strCadena = "SELECT id_tipo FROM linea_sub WHERE codtip='" & in_codt & "' and id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   get_sublinea = rst("id_tipo")
Else
    get_sublinea = "0"
End If

End Function

Private Sub Form_Activate()
KEY_RUC = Trim(Me.txtruc.Text)
'Me.lblempresa.Caption = get_persona(Trim(Me.txtruc.Text))
'Me.lbldireccion.Caption = get_direccion(Trim(Me.txtruc.Text))
End Sub

Private Sub Form_Load()

Skin1.LoadSkin App.Path & "\Skins\BS.skn"
Skin1.ApplySkin Me.hWnd




CenterForm Me

End Sub


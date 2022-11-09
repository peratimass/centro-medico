VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportesGenerales 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   BeginProperty Font 
      Name            =   "Bahnschrift SemiCondensed"
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
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_movimientos 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CON MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1020
      Left            =   3360
      TabIndex        =   84
      Top             =   4680
      Width           =   1515
   End
   Begin VB.OptionButton Opt_603 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "PRODUCTO POR VENDEDOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   83
      Top             =   8160
      Width           =   2775
   End
   Begin VB.CheckBox chk_cobertura 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "COVERTURA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   660
      Left            =   10080
      TabIndex        =   82
      Top             =   7800
      Width           =   1515
   End
   Begin VB.OptionButton opt_602 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTA POR PRODUCTO"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   81
      Top             =   7800
      Width           =   2775
   End
   Begin VB.OptionButton opt_504 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "INGRESOS Y SALIDAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   79
      Top             =   5760
      Width           =   2775
   End
   Begin VB.OptionButton opt_206 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "SALDO CLIENTE CREDITO"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   78
      Top             =   6480
      Width           =   2775
   End
   Begin VB.OptionButton opt_601 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "DESEMPEÑO POR VENDEDOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   76
      Top             =   7440
      Width           =   2775
   End
   Begin VB.OptionButton opt_503 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "ORDENES DE RECEPCION"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   75
      Top             =   5400
      Width           =   2775
   End
   Begin VB.OptionButton opt_502 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "ORDENES DE COMPRAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   74
      Top             =   5040
      Width           =   2775
   End
   Begin VB.OptionButton opt_501 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO DE COMPRAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   73
      Top             =   4680
      Width           =   2775
   End
   Begin VB.OptionButton opt_110 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "QUIEBRE DE STOCK"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   72
      Top             =   3840
      Width           =   2775
   End
   Begin VB.OptionButton opt_109 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "RESUMEN- MARGEN BRUTO"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   71
      Top             =   3480
      Width           =   2775
   End
   Begin VB.OptionButton opt_408 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LIQUIDACION DE VENTAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   69
      Top             =   3120
      Width           =   2775
   End
   Begin VB.OptionButton opt_106 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "PRODUCTO SERIES VENDIDAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   68
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton opt_407 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "BONIFICACIONES Y/OBSEQUIOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   67
      Top             =   2760
      Width           =   2775
   End
   Begin VB.OptionButton opt_406 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS PRODUCTO [ COBERTURAS ]"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   66
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton opt_405 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS POR VENDEDOR ARTICULO"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   65
      Top             =   2040
      Width           =   2775
   End
   Begin VB.OptionButton opt_404 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS POR VENDEDOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   64
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton opt_403 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS POR CATEGORIA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   63
      Top             =   1320
      Width           =   2775
   End
   Begin VB.OptionButton opt_402 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS X CLIENTE"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   62
      Top             =   960
      Width           =   2775
   End
   Begin VB.OptionButton opt_401 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VENTAS POR PRODUCTO"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7200
      TabIndex        =   61
      Top             =   600
      Width           =   2775
   End
   Begin VB.OptionButton opt_304 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "UTILIDADES X PRODUCTO Y LINEA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   59
      Top             =   8520
      Width           =   2775
   End
   Begin VB.OptionButton opt_303 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "UTILIDADES X LINEA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   58
      Top             =   8160
      Width           =   2775
   End
   Begin VB.OptionButton opt_301 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "UTILIDADES POR FECHA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   57
      Top             =   7440
      Width           =   2775
   End
   Begin VB.OptionButton opt_205 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CLIENTES POR VENDEDOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   55
      Top             =   6120
      Width           =   2775
   End
   Begin VB.OptionButton opt_302 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CLIENTES CON  [ + ] UTILIDAD"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   54
      Top             =   7800
      Width           =   2775
   End
   Begin VB.OptionButton opt_204 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CLIENTES CON  [ + ] VENTAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   53
      Top             =   5760
      Width           =   2775
   End
   Begin VB.OptionButton opt_203 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO DE TRANSPORTISTAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   52
      Top             =   5400
      Width           =   2775
   End
   Begin VB.OptionButton opt_202 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO DE PROVEEDORES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   51
      Top             =   5040
      Width           =   2775
   End
   Begin VB.OptionButton opt_201 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO DE CLIENTES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   50
      Top             =   4680
      Width           =   2775
   End
   Begin VB.OptionButton opt_108 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "PRODUCTOS PENDIENTE ENTREGA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   48
      Top             =   3120
      Width           =   2775
   End
   Begin VB.OptionButton opt_107 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "PRODUCTOS CON (+) ROTACION"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   47
      Top             =   2760
      Width           =   2775
   End
   Begin VB.OptionButton opt_105 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "PRODUCTO SERIES DISPONIBLES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   46
      Top             =   2040
      Width           =   2775
   End
   Begin VB.OptionButton opt_104 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "KARDEX"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   45
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton opt_103 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "VALORIZADO STOCK  [NOW]"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   44
      Top             =   1320
      Width           =   2775
   End
   Begin VB.OptionButton opt_102 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CATALOGO DE PRECIOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   43
      Top             =   960
      Width           =   2775
   End
   Begin VB.CheckBox chk_todos 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "TODOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   4320
      TabIndex        =   42
      Top             =   600
      Value           =   1  'Checked
      Width           =   920
   End
   Begin VB.CheckBox chk_stock 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   3360
      TabIndex        =   41
      Top             =   600
      Width           =   920
   End
   Begin VB.OptionButton opt_101 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "LISTADO DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   480
      TabIndex        =   40
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame frmParametros 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   12600
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.TextBox txtVendedor 
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
         Left            =   6480
         TabIndex        =   38
         Top             =   4560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3840
         TabIndex        =   32
         Top             =   6720
         Width           =   2775
         Begin VB.CheckBox chk_resumido 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "RESUMIDO"
            ForeColor       =   &H00800000&
            Height          =   320
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chk_detallado 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "DETALLADO"
            ForeColor       =   &H00800000&
            Height          =   320
            Left            =   1320
            TabIndex        =   33
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.CheckBox chk_sucursal 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SUCURSAL"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chk_vendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "VENDEDOR"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   12
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CheckBox chk_clasificacion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "FAMILIA"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chk_subfamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SUB FAMILIA "
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   10
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CheckBox chk_proveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PROVEEDOR"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   9
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CheckBox chk_marca 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "MARCA"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   8
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CheckBox chk_producto 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTO"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   7
         Top             =   4080
         Width           =   1815
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
         Left            =   6480
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3840
         TabIndex        =   5
         Top             =   6120
         Width           =   2655
         Begin VB.CheckBox chk_ordenar_descripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "DESCRIPCION"
            ForeColor       =   &H00800000&
            Height          =   320
            Left            =   1320
            TabIndex        =   30
            Top             =   0
            Width           =   1335
         End
         Begin VB.CheckBox chk_ordenar_codigo 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CODIGO"
            ForeColor       =   &H00800000&
            Height          =   320
            Left            =   120
            TabIndex        =   29
            Top             =   0
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.CheckBox chk_modelo 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "MODELO"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   4
         Top             =   3060
         Width           =   1815
      End
      Begin VB.TextBox TxtCliente 
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
         Left            =   6480
         TabIndex        =   3
         Top             =   5040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chk_cliente 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTE"
         ForeColor       =   &H00800000&
         Height          =   320
         Left            =   360
         TabIndex        =   2
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox TxtProveedor 
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
         Left            =   6480
         TabIndex        =   1
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiBold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   136839169
         CurrentDate     =   43096
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   2400
         TabIndex        =   15
         Top             =   1560
         Width           =   4000
         _ExtentX        =   7064
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
         Left            =   4680
         TabIndex        =   16
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiBold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   136839169
         CurrentDate     =   43096
      End
      Begin MSDataListLib.DataCombo DtcVendedor 
         Height          =   330
         Left            =   2400
         TabIndex        =   17
         Top             =   4560
         Visible         =   0   'False
         Width           =   4000
         _ExtentX        =   7064
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
         TabIndex        =   18
         Top             =   2040
         Width           =   4000
         _ExtentX        =   7064
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
      Begin MSDataListLib.DataCombo DtcSubfamilia 
         Height          =   330
         Left            =   2400
         TabIndex        =   19
         Top             =   2520
         Width           =   4000
         _ExtentX        =   7064
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
         TabIndex        =   20
         Top             =   5520
         Visible         =   0   'False
         Width           =   4000
         _ExtentX        =   7064
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
         TabIndex        =   21
         Top             =   3600
         Width           =   4000
         _ExtentX        =   7064
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
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   330
         Left            =   2400
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
         Width           =   4000
         _ExtentX        =   7064
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
         TabIndex        =   23
         Top             =   3060
         Width           =   4000
         _ExtentX        =   7064
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
      Begin MSDataListLib.DataCombo DtcCliente 
         Height          =   330
         Left            =   2400
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   4000
         _ExtentX        =   7064
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
      Begin VitekeySoft.ChameleonBtn cmd_reporte 
         Height          =   615
         Left            =   2520
         TabIndex        =   35
         Top             =   8160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "GENERAR REPORTE"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
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
         MICON           =   "frmReportesGenerales.frx":0000
         PICN            =   "frmReportesGenerales.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_cerrar 
         Height          =   615
         Left            =   5040
         TabIndex        =   36
         Top             =   8160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
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
         MICON           =   "frmReportesGenerales.frx":25ED
         PICN            =   "frmReportesGenerales.frx":2609
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
         Left            =   2520
         TabIndex        =   37
         Top             =   7920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "TIPO REPORTE"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2400
         TabIndex        =   31
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PARAMETROS DE BUSQUEDA"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiBold SemiConden"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2400
         TabIndex        =   28
         Top             =   120
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "RANGO FECHAS  :"
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "ORDENAR POR:"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2400
         TabIndex        =   26
         Top             =   6120
         Width           =   1335
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
         Left            =   4080
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         Height          =   8880
         Left            =   0
         Top             =   0
         Width           =   7425
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdUpdateCosto 
      Height          =   1455
      Left            =   3360
      TabIndex        =   80
      Top             =   7440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   2566
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
      MICON           =   "frmReportesGenerales.frx":29F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "VENDEDORES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   3
      Left            =   6660
      TabIndex        =   77
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Index           =   4
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRAS - INGRESOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   2
      Left            =   6690
      TabIndex        =   70
      Top             =   4320
      Width           =   1605
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "VENTAS-SALIDAS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   6720
      TabIndex        =   60
      Top             =   195
      Width           =   1275
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "UTILIDADES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   285
      TabIndex        =   56
      Top             =   7080
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTES Y PROVEEDORES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   315
      TabIndex        =   49
      Top             =   4305
      Width           =   1905
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS Y SERVICIOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold SemiConden"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   270
      TabIndex        =   39
      Top             =   200
      Width           =   1755
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   375
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmReportesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub chk_clasificacion_Click()
If Me.chk_clasificacion.Value = 1 Then
    strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(DtcLinea)

End If

End Sub

Private Sub chk_cliente_Click()

If Me.chk_cliente.Value = 1 Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_cliente='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcCliente)
    Me.TxtCliente.Visible = True
    Me.DtcCliente.Visible = True
End If


End Sub

Private Sub chk_detallado_Click()
If Me.chk_detallado.Value = 1 Then
   Me.chk_resumido.Value = 0
Else
   Me.chk_resumido.Value = 1
End If
End Sub

Private Sub chk_marca_Click()
If Me.chk_marca.Value = 1 Then
        strCadena = "SELECT id_marca as Codigo,descripcion as Descripcion FROM marca WHERE  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcMarca)
End If


End Sub

Private Sub chk_modelo_Click()

If Me.chk_modelo.Value = 1 Then
    
  
  strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM linea_modelo WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtcModelo.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcModelo)
  
End If

End Sub

Private Sub chk_ordenar_codigo_Click()
If Me.chk_ordenar_codigo.Value = 1 Then
   Me.chk_ordenar_descripcion.Value = 0
Else
   Me.chk_ordenar_descripcion.Value = 1
End If
End Sub

Private Sub chk_producto_Click()

If Me.chk_producto.Value = 1 Then
  strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProducto)
  Me.DtcProducto.Visible = True
  Me.Txtproducto.Visible = True
  
End If

End Sub

Private Sub chk_proveedor_Click()

If Me.chk_proveedor.Value = 1 Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_proveedor='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedor)
    Me.DtcProveedor.Visible = True
    Me.TxtProveedor.Visible = True
  
End If


End Sub

Private Sub chk_resumido_Click()
If Me.chk_resumido.Value = 1 Then
    Me.chk_detallado.Value = 0
Else
   Me.chk_detallado.Value = 1
End If
End Sub

Private Sub chk_stock_Click()
If Me.chk_stock.Value = 1 Then
   Me.chk_todos.Value = 0
Else
    Me.chk_todos.Value = 1
End If
End Sub

Private Sub chk_subfamilia_Click()
If Me.chk_subfamilia.Value = 1 Then
   strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_usu='" & KEY_RUC & "' ORDER BY descripcion"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcSubfamilia)
End If
End Sub

Private Sub chk_sucursal_Click()

If Me.chk_sucursal.Value = 1 Then
      strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and  ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
End If

End Sub

Private Sub chk_todos_Click()
If Me.chk_todos.Value = 1 Then
   Me.chk_stock.Value = 0
Else
   Me.chk_stock.Value = 1
End If
End Sub

Private Sub chk_vendedor_Click()
If Me.chk_vendedor.Value = 1 Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcVendedor)
    Me.DtcVendedor.Visible = True
    Me.txtVendedor.Visible = True
  
End If
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_reporte_Click()
Dim in_nombre_reporte As String
Dim cam3(0 To 5, 1 To 5)  As String
'Filtros
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
    
    
    If Me.chk_subfamilia.Value = 1 Then
       in_sublinea = Me.DtcSubfamilia.BoundText
    Else
       in_sublinea = ""
    End If
    
    If Me.chk_modelo.Value = 1 Then
       in_modelo = Me.DtcModelo.BoundText
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
    
    If Me.chk_cliente.Value = 1 Then
        in_cliente = Me.DtcCliente.BoundText
    Else
        in_cliente = ""
    End If
    
    If Me.chk_proveedor.Value = 1 Then
        in_proveedor = Me.DtcProveedor.BoundText
    Else
        in_proveedor = ""
    End If
    
    If Me.chk_vendedor.Value = 1 Then
        in_vendedor = Me.DtcVendedor.BoundText
    Else
        in_vendedor = ""
    End If
    
    
    If Me.chk_resumido.Value = 1 Then
       in_tipo_reporte = 1
    Else
       in_tipo_reporte = 2
    End If
    
    If Me.chk_ordenar_codigo.Value = 1 Then
       in_ordenamiento = 1
    Else
       in_ordenamiento = 2
    End If
    
    
    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_almacen
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    
    
    
    
    If Me.opt_204.Value = True Then
        cam3(5, 2) = "CLIENTES CON [+] VENTAS"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('1','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If in_tipo_reporte = 1 Then
           in_nombre_reporte = "RptClientemasventa"
        Else
           in_nombre_reporte = "Rpt204"
        End If
        Ans = ShowMultiReport(rst, in_nombre_reporte, param, App.Path + "\Reportes\")
        Exit Sub
    End If

 If Me.opt_504.Value = True Then
        cam3(5, 2) = "INGRESAS Y SALIDAS DE PRODUCTOS"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('2','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        
        Ans = ShowMultiReport(rst, "Rpt504", param, App.Path + "\Reportes\")
        Exit Sub
    End If
 
    If Me.opt_302.Value = True Then
        cam3(5, 2) = "CLIENTES CON [+] UTILIDAD"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('3','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If in_tipo_reporte = 1 Then
           in_nombre_reporte = "RptClienteUtilidad"
        Else
           in_nombre_reporte = "Rpt302"
        End If
        Ans = ShowMultiReport(rst, in_nombre_reporte, param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    
    
    If Me.opt_101.Value = True Then
        cam3(5, 2) = "LISTADO DE PRODUCTOS."
        param = cam3()
        If Me.chk_stock.Value = 1 Then
           in_operacion = 5
        Else
           in_operacion = 4
        End If
        
        
        strCadena = "CALL ADM_reportes_generales_v3('" & in_operacion & "','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        in_nombre_reporte = "RptProducto1"
        Ans = ShowMultiReport(rst, in_nombre_reporte, param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    If Me.opt_102.Value = True Then
        cam3(5, 2) = "CATALOGO DE PRODUCTOS."
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('6','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptProducto_precios", param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    If Me.opt_103.Value = True Then
        cam3(5, 2) = "VALORIZADO PRODUCTOS [STOCK]."
        param = cam3()
        in_operacion = 10
        strCadena = "CALL ADM_reportes_generales_v3('" & in_operacion & "','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        in_nombre_reporte = "RptProductoValorizado"
        Ans = ShowMultiReport(rst, in_nombre_reporte, param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If Me.opt_104.Value = True Then
        
        cam3(5, 2) = "KARDEX VALORIZADO [ AL :" + Format(Me.DtpInicio.Value, "dd-mm-YYYY") + " ]"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('11','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptInventarioValorizado", param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    
    If Me.opt_201.Value = True Then
        If Me.chk_movimientos.Value = 1 Then
            in_tipo_reporte = 1
        Else
            in_tipo_reporte = 0
        End If
        
        cam3(5, 2) = "LISTADO DE CLIENTES [ AL :" + Format(Me.DtpInicio.Value, "dd-mm-YYYY") + " ]"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('13','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptCliente", param, App.Path + "\Reportes\", "Reporte de Clientes")
        Exit Sub
    End If
    If Me.opt_202.Value = True Then
        If Me.chk_movimientos.Value = 1 Then
            in_tipo_reporte = 1
        Else
            in_tipo_reporte = 0
        End If
        cam3(5, 2) = "LISTADO DE PROVEEDORES [ AL :" + Format(Me.DtpInicio.Value, "dd-mm-YYYY") + " ]"
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('14','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptCliente", param, App.Path + "\Reportes\", "Reporte de Proveedores")
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
   
   If Me.opt_601.Value = True Then
        cam3(5, 2) = "ESTADISTICA DESEMPEÑO X VENDEDOR"
        param = cam3()
        Call persona_rendimiento(in_producto, in_linea, in_sublinea, in_modelo, in_marca, in_vendedor, in_proveedor)
       strCadena = "call ADM_reportes_generales_v3('9','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptDesempenio", param, App.Path + "\Reportes\")
        Exit Sub
   End If
   
    
    If Me.opt_602.Value = True Then
        If Me.chk_cobertura.Value = 1 Then
            Call put_cobertura_vendedor(in_proveedor)
        Else
            strCadena = "CALL ADM_get_cobertura('0','','','','','0','','" & KEY_USUARIO & "','" & KEY_RUC & "') "
            CnBd.Execute (strCadena)
        End If
        cam3(5, 2) = "VENTAS POR PRODUCTO."
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('7','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptVentaVendedor", param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    
    
    
    
    If Me.Opt_603.Value = True Then
        If Me.chk_cobertura.Value = 1 Then
            Call put_cobertura_vendedor(in_proveedor)
        Else
            strCadena = "CALL ADM_get_cobertura('0','','','','','0','','" & KEY_USUARIO & "','" & KEY_RUC & "') "
            CnBd.Execute (strCadena)
        End If
        cam3(5, 2) = "VENTAS POR VENDEDOR."
        param = cam3()
        strCadena = "CALL ADM_reportes_generales_v3('12','" & in_tipo_reporte & "','" & in_ordenamiento & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & in_cliente & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_marca & "','" & in_proveedor & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptVentaVendedor2", param, App.Path + "\Reportes\")
        Exit Sub
    End If
    
    
    
    
   
   Exit Sub
   
   
   
   
   
  
    
    
    
    MsgBox "DEBE SELECCIONAR UNA OPCION.", vbInformation, KEY_EMPRESA
    
    
    
    
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


Private Sub put_cobertura_vendedor(ByVal in_proveedor As String)
On Error GoTo salir
strCadena = "CALL ADM_get_cobertura('0','','','','','0','','" & KEY_USUARIO & "','" & KEY_RUC & "') "
CnBd.Execute (strCadena)
          
strCadena = "call ADM_get_cobertura('3','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','','','0','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prg_reporte.Min = 1
   Me.prg_reporte.Max = rst.RecordCount
   For i = 1 To rst.RecordCount
       strCadena = "call ADM_get_cobertura('1','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & rst("id_vendedor") & "','" & in_proveedor & "','0','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          rstK.MoveFirst
          
          For j = 1 To rstK.RecordCount
               strCadena = "CALL ADM_get_cobertura('2','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & rst("id_vendedor") & "','" & in_proveedor & "','0','" & rstK("id_producto") & "','" & KEY_USUARIO & "','" & KEY_RUC & "') "
               CnBd.Execute (strCadena)
               rstK.MoveNext
               
               Me.cmd_reporte.Caption = "[" & str(i) & Space(2) & str(rst.RecordCount) & "]" & Space(2) + str(j) & Space(2) & str(rstK.RecordCount)
               DoEvents
          Next j
        End If
       Me.prg_reporte.Value = i
       DoEvents
       rst.MoveNext
   Next i
End If
Exit Sub
salir:
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
            
            If Round(rstIN(0), 6) <> Round(rst("precio_costo"), 6) Then
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

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA


   
  

  
  


End Sub

Private Sub txtcliente_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo like '%" & Trim(Me.TxtCliente.Text) & "%' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcCliente)
End Sub

Private Sub TxtProducto_Change()
    strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto WHERE nombre_prod LIKE '%" & Trim(Me.Txtproducto.Text) & "%'  and  ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProducto)
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub TxtProveedor_Change()
    
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.TxtProveedor.Text) & "%' and  id_proveedor='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedor)

End Sub

Private Sub txtVendedor_Change()
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtVendedor.Text) & "%' and  id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo "
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcVendedor)
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteIngresos 
   BorderStyle     =   0  'None
   Caption         =   "REGISTRO DE INGRESOS"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SstKardex 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "INGRESOS"
      TabPicture(0)   =   "FrmReporteIngresos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(2)=   "LblCantidad"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "cmdcerrar"
      Tab(0).Control(5)=   "cmdprocesar"
      Tab(0).Control(6)=   "DtcProducto"
      Tab(0).Control(7)=   "DtpHasta"
      Tab(0).Control(8)=   "DtpDesde"
      Tab(0).Control(9)=   "DtcAlmacen"
      Tab(0).Control(10)=   "chkAlmacen"
      Tab(0).Control(11)=   "ChkProducto"
      Tab(0).Control(12)=   "txtBuscar"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "SALIDAS DE MERCADERIA"
      TabPicture(1)   =   "FrmReporteIngresos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DtcCliente"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DtcTipodoc"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdcerrar_salida"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdProcesar_salidas"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DtcProducto_salida"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DtpFin_salida"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "DtpInicio_salida"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "DtcSucursal_salidas"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chk_producto_salida"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chk_almacen_salida"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chk_tipodocumento"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Check1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TxtBuscarCliente"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin VB.TextBox TxtBuscarCliente 
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
         Left            =   7200
         TabIndex        =   26
         Top             =   3870
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
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
         Height          =   300
         Left            =   900
         TabIndex        =   24
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chk_tipodocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TIPO DOC :"
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
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox txtBuscar 
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
         Left            =   -67800
         TabIndex        =   21
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox chk_almacen_salida 
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
         Height          =   300
         Left            =   885
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chk_producto_salida 
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
         Height          =   300
         Left            =   900
         TabIndex        =   11
         Top             =   2480
         Width           =   1215
      End
      Begin VB.CheckBox ChkProducto 
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
         Height          =   300
         Left            =   -74160
         TabIndex        =   7
         Top             =   2550
         Width           =   1215
      End
      Begin VB.CheckBox chkAlmacen 
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
         Height          =   300
         Left            =   -74175
         TabIndex        =   1
         Top             =   1890
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   -72810
         TabIndex        =   2
         Top             =   1860
         Width           =   4935
         _ExtentX        =   8705
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
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   -74160
         TabIndex        =   3
         Top             =   915
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
         Format          =   122945537
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   -72000
         TabIndex        =   4
         Top             =   915
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
         Format          =   122945537
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   330
         Left            =   -72795
         TabIndex        =   8
         Top             =   2520
         Width           =   4935
         _ExtentX        =   8705
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   975
         Left            =   -65040
         TabIndex        =   9
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
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
         MICON           =   "FrmReporteIngresos.frx":0038
         PICN            =   "FrmReporteIngresos.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   975
         Left            =   -65040
         TabIndex        =   10
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
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
         MICON           =   "FrmReporteIngresos.frx":2625
         PICN            =   "FrmReporteIngresos.frx":2641
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcSucursal_salidas 
         Height          =   330
         Left            =   2130
         TabIndex        =   13
         Top             =   1815
         Width           =   4935
         _ExtentX        =   8705
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
      Begin MSComCtl2.DTPicker DtpInicio_salida 
         Height          =   315
         Left            =   900
         TabIndex        =   14
         Top             =   870
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
         Format          =   122945537
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpFin_salida 
         Height          =   315
         Left            =   3180
         TabIndex        =   15
         Top             =   870
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
         Format          =   122945537
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcProducto_salida 
         Height          =   330
         Left            =   2145
         TabIndex        =   16
         Top             =   2475
         Width           =   4935
         _ExtentX        =   8705
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar_salidas 
         Height          =   975
         Left            =   9780
         TabIndex        =   19
         Top             =   1035
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
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
         MICON           =   "FrmReporteIngresos.frx":2A31
         PICN            =   "FrmReporteIngresos.frx":2A4D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar_salida 
         Height          =   975
         Left            =   9780
         TabIndex        =   20
         Top             =   2115
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
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
         MICON           =   "FrmReporteIngresos.frx":501E
         PICN            =   "FrmReporteIngresos.frx":503A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcTipodoc 
         Height          =   330
         Left            =   2145
         TabIndex        =   23
         Top             =   3120
         Width           =   4935
         _ExtentX        =   8705
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
      Begin MSDataListLib.DataCombo DtcCliente 
         Height          =   330
         Left            =   2145
         TabIndex        =   25
         Top             =   3840
         Width           =   4935
         _ExtentX        =   8705
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
      Begin VB.Label Label3 
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
         Left            =   2595
         TabIndex        =   18
         Top             =   915
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RANGO DE FECHAS"
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
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RANGO DE FECHAS"
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
         Left            =   -74220
         TabIndex        =   6
         Top             =   525
         Width           =   1305
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
         Left            =   -72465
         TabIndex        =   5
         Top             =   960
         Width           =   225
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0C0&
         Height          =   4215
         Left            =   -74760
         Top             =   360
         Width           =   10815
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   -74400
         Top             =   1560
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0C0&
         Height          =   4215
         Left            =   240
         Top             =   360
         Width           =   10815
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6720
      Top             =   2655
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
            Picture         =   "FrmReporteIngresos.frx":542A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":587E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":5B9E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":5FF2
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":6446
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":6766
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":6BBA
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":6D16
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":716A
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":7486
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":7D62
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":8082
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteIngresos.frx":83A2
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "FrmReporteIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAlmacen_Click()
If Me.ChkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If
End Sub

Private Sub ChkProducto_Click()
If Me.chkProducto.Value = 1 Then
    Me.DtcProducto.Enabled = True
Else
    Me.DtcProducto.Enabled = False
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdcerrar_salida_Click()
Unload Me
End Sub

Private Sub cmdProcesar_Click()
   Dim in_criterio As String
   If Me.chkProducto.Value = 1 Then
        in_criterio = " and id_producto='" & Me.DtcProducto.BoundText & "'"
   Else
        in_criterio = ""
   End If
   
   
   If KEY_RUC = "20128836251" Then
        strCadena = "SELECT id_orden,fecha_registro,numero,id_proveedor,nombre_completo,id_producto,nombre_prod,cantidad,precio,total FROM view_ingresos WHERE  fecha_registro>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_registro<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'  ORDER BY fecha_registro asc"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptIngreso_recepciones", , App.Path + "\Reportes\")
        Exit Sub
   Else
   
   If Me.ChkAlmacen.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpDesde.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpHasta.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_compra,fecha_emision,id_producto,nombre_prod,cantidad,nombre_completo,comprobante,moneda,c_unitario,otros,dsto_soles,total FROM view_ingreso_producto WHERE  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'" & in_criterio & " ORDER BY fecha_emision DESC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpDesde.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpHasta.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_compra,fecha_emision,id_producto,nombre_prod,cantidad,nombre_completo,comprobante,moneda,c_unitario,otros,dsto_soles,total FROM view_ingreso_producto WHERE  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' " & in_criterio & " ORDER BY fecha_emision DESC"
   End If
   End If
   
   
   
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptIngreso_producto", , App.Path + "\Reportes\")
   
   
End Sub

Private Sub cmdProcesar_salidas_Click()
 Dim cam3(0 To 2, 1 To 2)  As String
 cam3(0, 1) = "inicial"
 cam3(1, 1) = "final"
 cam3(2, 1) = "almacen"
 cam3(0, 2) = Format(Me.DtpInicio_salida.Value, "dd-mm-YYYY")
 cam3(1, 2) = Format(Me.DtpFin_salida.Value, "dd-mm-YYYY")
 cam3(2, 2) = ""
 param = cam3()
   
       If Me.Check1.Value = 1 Then
          cam3(2, 2) = Me.dtcCliente.BoundText & Space(2) & get_persona(Me.dtcCliente.BoundText)
          in_proveedor = Me.dtcCliente.BoundText
       Else
         cam3(2, 2) = "TODOS LAS SALIDAS"
         in_proveedor = ""
       End If
        
           
        
         
        
        

   param = cam3()
   
   strCadena = "call ADM_reportes_generales('2','" & Format(Me.DtpInicio_salida.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin_salida.Value, "YYYY-mm-dd") & "','" & in_proveedor & "','" & in_producto & "','" & in_alm & "','" & in_linea & "','" & in_subfamilia & "','" & in_modelo & "','" & in_marca & "','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "Rpt101", param, App.Path + "\Reportes\")
   Exit Sub
   
   
   
   
   
   
   Dim in_criterio As String
   If Me.chkProducto.Value = 1 Then
        in_criterio = " and id_producto='" & Me.DtcProducto.BoundText & "'"
   Else
        in_criterio = ""
   End If
   
   
   
   If Me.chk_tipodocumento.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpInicio_salida.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin_salida.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_venta,fecha_emision,id_producto,nombre_prod,cantidad,nombre_completo,comprobante,moneda,c_unitario,otros,dsto_soles,total FROM view_salida_producto WHERE  fecha_emision>='" & Format(Me.DtpInicio_salida.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_salida.Value, "YYYY-mm-dd") & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and    ruc='" & KEY_RUC & "'" & in_criterio & " ORDER BY fecha_emision DESC"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptIngreso_producto", , App.Path + "\Reportes\")
        Exit Sub
   End If
   
   
   
   If KEY_RUC = "20128836251" Then
        strCadena = "SELECT id_orden,fecha_registro,numero,id_proveedor,nombre_completo,id_producto,nombre_prod,cantidad FROM view_ordenes_salida WHERE  fecha_registro>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_registro<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'  ORDER BY fecha_registro asc"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptIngreso_recepciones", , App.Path + "\Reportes\")
        Exit Sub
   Else
   
   
   
   If Me.chk_almacen_salida.Value = 1 Then
        strCadena = "SELECT '" & Format(Me.DtpDesde.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpHasta.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_venta,fecha_emision,id_producto,nombre_prod,cantidad,nombre_completo,comprobante,moneda,c_unitario,otros,dsto_soles,total FROM view_salida_producto WHERE  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'" & in_criterio & " ORDER BY fecha_emision DESC"
   Else
        strCadena = "SELECT '" & Format(Me.DtpDesde.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpHasta.Value, "dd-mm-YYYY") & "','" & Me.DtcAlmacen.Text & "',id_venta,fecha_emision,id_producto,nombre_prod,cantidad,nombre_completo,comprobante,moneda,c_unitario,otros,dsto_soles,total FROM view_salida_producto WHERE  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' " & in_criterio & " ORDER BY fecha_emision DESC"
   End If
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptIngreso_producto", , App.Path + "\Reportes\")
   
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA

Me.DtpInicio_salida.Value = KEY_FECHA
Me.DtpFin_salida.Value = KEY_FECHA


strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSucursal_salidas)
Me.DtcSucursal_salidas.BoundText = KEY_ALM
Me.DtcSucursal_salidas.Enabled = False


strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
Me.DtcProducto.Enabled = False


strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes  ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoDoc)





strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto_salida)
Me.DtcProducto.Enabled = False

End Sub


Private Sub txtBuscar_Change()

If Me.chkProducto.Value = 1 Then
strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE nombre_prod LIKE '%" & Trim(Me.txtBuscar.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
End If

End Sub

Private Sub TxtBuscarCliente_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.TxtBuscarCliente.Text) & "%'  and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.dtcCliente)

End Sub

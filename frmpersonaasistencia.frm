VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmpersonaasistencia 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtprecio_hora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   12360
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtanio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   9600
      TabIndex        =   32
      Top             =   540
      Width           =   1215
   End
   Begin VB.TextBox txtbuscar_trabajador 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   12360
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame frmpermiso 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   11175
      Begin VB.TextBox txtid_permiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8520
         TabIndex        =   37
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkturno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TURNO :"
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
         Height          =   375
         Left            =   1050
         TabIndex        =   36
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtnumerodias 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2640
         TabIndex        =   25
         Text            =   "1"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtautorizadopor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Left            =   2640
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3960
         Width           =   4335
      End
      Begin VB.OptionButton optnoremunerado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO REMUNERADO"
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
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   3480
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optremunerado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "REMUNERADO"
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
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   3480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpinicio 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Format          =   121176065
         CurrentDate     =   42236
      End
      Begin VB.TextBox txtdetallepermiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   675
         Left            =   2640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtbuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7080
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DtcTrabajador 
         Height          =   330
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   405
         Left            =   4200
         TabIndex        =   19
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "PROCESAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
         MICON           =   "frmpersonaasistencia.frx":0000
         PICN            =   "frmpersonaasistencia.frx":001C
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
         Left            =   5640
         TabIndex        =   20
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "CERRAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
         MICON           =   "frmpersonaasistencia.frx":2601
         PICN            =   "frmpersonaasistencia.frx":261D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmddias 
         Height          =   350
         Left            =   4080
         TabIndex        =   23
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
         MICON           =   "frmpersonaasistencia.frx":5632
         PICN            =   "frmpersonaasistencia.frx":564E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdimprimir 
         Height          =   405
         Left            =   2760
         TabIndex        =   34
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "IMPRIMIR"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
         MICON           =   "frmpersonaasistencia.frx":83E4
         PICN            =   "frmpersonaasistencia.frx":8400
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcTurno 
         Height          =   330
         Left            =   2640
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label lbldias 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº DIAS :"
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
         TabIndex        =   24
         Top             =   2400
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTORIZADO POR :"
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
         Left            =   660
         TabIndex        =   21
         Top             =   4005
         Width           =   1485
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERMISO REMUNERADO :"
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
         Left            =   180
         TabIndex        =   16
         Top             =   3480
         Width           =   1965
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIO :"
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
         Left            =   1020
         TabIndex        =   14
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DEL PERMISO :"
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
         Left            =   330
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Left            =   1020
         TabIndex        =   9
         Top             =   405
         Width           =   1125
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevopermiso 
      Height          =   855
      Left            =   13560
      TabIndex        =   5
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "frmpersonaasistencia.frx":848D
      PICN            =   "frmpersonaasistencia.frx":84A9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox TxtApellido 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   10821
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
   Begin VitekeySoft.ChameleonBtn cmdreportepersonal 
      Height          =   855
      Left            =   13560
      TabIndex        =   6
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "INDIVIDUAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "frmpersonaasistencia.frx":AD93
      PICN            =   "frmpersonaasistencia.frx":ADAF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdreportegeneral 
      Height          =   855
      Left            =   13560
      TabIndex        =   7
      Top             =   2880
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "GENERAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "frmpersonaasistencia.frx":DFB3
      PICN            =   "frmpersonaasistencia.frx":DFCF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTrabajadorReporte 
      Height          =   330
      Left            =   7920
      TabIndex        =   28
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   855
      Left            =   13560
      TabIndex        =   29
      Top             =   3840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "CERRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "frmpersonaasistencia.frx":111D3
      PICN            =   "frmpersonaasistencia.frx":111EF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcMes 
      Height          =   330
      Left            =   7920
      TabIndex        =   31
      Top             =   540
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MES :"
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
      Left            =   7320
      TabIndex        =   30
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE TRABAJADOR :"
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
      Left            =   5880
      TabIndex        =   26
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE TRABAJADOR :"
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
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   14655
   End
End
Attribute VB_Name = "frmpersonaasistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChameleonBtn1_Click()

End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub ChameleonBtn3_Click()
Unload Me
End Sub

Private Sub chkturno_Click()
If Me.chkturno.Value = 1 Then
    Me.DtcTurno.Visible = True
Else
    Me.DtcTurno.Visible = False
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmddias_Click()
If Me.txtnumerodias.Visible = False Then
    Me.lbldias.Visible = True
    Me.txtnumerodias.Visible = True
    Call Resalta(Me.txtnumerodias)
Else
     Me.lbldias.Visible = False
    Me.txtnumerodias.Visible = False
    
End If
End Sub

Private Sub cmdimprimir_Click()
strCadena = "SELECT p.dni,pp.nombre_completo,pp.direccion,p.fecha_permiso,p.fecha_permiso,t.descripcion,p.detalle,p.autorizado,p.usuario FROM persona_permiso p,persona pp,turno_hc t WHERE p.dni=pp.dni and p.id_turno=t.id_turno and p.ruc=t.ruc and p.id_permiso='" & Val(Me.txtid_permiso.Text) & "' and p.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "rpt_permiso", , App.Path + "\Reportes\")
End Sub

Private Sub cmdnuevopermiso_Click()

Me.frmpermiso.Visible = True
Me.cmdimprimir.Enabled = False
Me.cmdprocesar.Enabled = True
End Sub

Private Sub cmdprocesar_Click()

Dim str_remunerado As String
If Me.optnoremunerado.Value = True Then
    str_remunerado = "si"
Else
    str_remunerado = "no"
End If

If Me.chkturno.Value = 1 Then
    str_turno = Me.DtcTurno.BoundText
Else
    str_turno = "-"
End If


If Me.txtnumerodias.Visible = True And Val(Me.txtnumerodias.Text) > 1 Then
    For i = 0 To Val(Me.txtnumerodias.Text) - 1
        strCadena = "INSERT INTO persona_permiso(`fecha_registro`,`dni`,`detalle`,`dni_save`,`fecha_permiso`,`remunerado`,`autorizado`,usuario,id_turno,`ruc`) " & _
        " VALUES(CURDATE(),'" & Me.DtcTrabajador.BoundText & "','" & Trim(Me.txtdetallepermiso.Text) & "','" & KEY_USUARIO & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & str_remunerado & "','" & Trim(Me.txtautorizadopor.Text) & "','" & KEY_VENDEDOR & "','" & str_turno & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        strCadena = "INSERT INTO persona_asistencia(dni,ruc,fecha,hora,hora_inicio,horas_trabajo,id_acceso,dni_save)VALUES('" & Me.DtcTrabajador.BoundText & "','" & KEY_RUC & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Time, "HH:mm:ss") & "','00:00','0','03','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
         
            
        
        Me.DtpInicio.Value = DateAdd("d", 1, Me.DtpInicio.Value)
    Next i
Else
    strCadena = "INSERT INTO persona_permiso(`fecha_registro`,`dni`,`detalle`,`dni_save`,`fecha_permiso`,`remunerado`,`autorizado`,usuario,id_turno,`ruc`) " & _
    " VALUES(CURDATE(),'" & Me.DtcTrabajador.BoundText & "','" & Trim(Me.txtdetallepermiso.Text) & "','" & KEY_USUARIO & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & str_remunerado & "','" & Trim(Me.txtautorizadopor.Text) & "','" & KEY_VENDEDOR & "','" & str_turno & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    
    strCadena = "INSERT INTO persona_asistencia(dni,ruc,fecha,hora,hora_inicio,horas_trabajo,id_acceso,dni_save)VALUES('" & Me.DtcTrabajador.BoundText & "','" & KEY_RUC & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Time, "HH:mm:ss") & "','00:00','0','03','" & KEY_USUARIO & "')"
    CnBd.Execute (strCadena)
     
            
End If
strCadena = "SELECT * FROM persona_permiso ORDER BY id_permiso DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtid_permiso.Text = rst("id_permiso")
End If


strCadena = "SELECT * FROM view_permiso WHERE ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfdPersona)
Me.cmdimprimir.Enabled = True
Me.cmdprocesar.Enabled = False
'Me.frmpermiso.Visible = False
End Sub

Private Sub cmdreportegeneral_Click()
Dim fecha_ini As String
Dim fecha_fin As String
Dim nhoras As Single
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
fecha_ini = Trim(Me.txtanio.Text) & "-" & Format(Me.dtcmes.BoundText, "00") & "-" & "01"
ultimo_dia = DateSerial(Me.txtanio.Text, Val(Me.dtcmes.BoundText) + 1, 0)
fecha_fin = Trim(Me.txtanio.Text) & "-" & Format(Me.dtcmes.BoundText, "00") & "-" & ultimo_dia
arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(2, 1) = "p_precio_hora"



arr(0, 2) = fecha_ini
arr(1, 2) = fecha_fin
arr(2, 2) = str(Me.txtprecio_hora.Text)

param = arr()


strCadena = "SELECT  dni,fecha,hora,horas_trabajo,id_acceso,nombre_completo FROM view_reporte_asistencia WHERE ruc='" & KEY_RUC & "' and fecha>='" & Format(fecha_ini, "YYYY-mm-dd") & "'  and fecha<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY id"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_asistencia", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdreportepersonal_Click()
Dim fecha_ini As String
Dim fecha_fin As String
Dim nhoras As Single
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
fecha_ini = Trim(Me.txtanio.Text) & "-" & Format(Me.dtcmes.BoundText, "00") & "-" & "01"
ultimo_dia = DateSerial(Me.txtanio.Text, Val(Me.dtcmes.BoundText) + 1, 0)
fecha_fin = Format(CVDate(Trim(Me.txtanio.Text) & "-" & Format(Me.dtcmes.BoundText, "00") & "-" & Format(Day(ultimo_dia), "00")), "YYYY-mm-dd")
arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(2, 1) = "p_precio_hora"


arr(0, 2) = fecha_ini
arr(1, 2) = fecha_fin
arr(2, 2) = str(Me.txtprecio_hora.Text)

param = arr()


strCadena = "SELECT  dni,fecha,hora,horas_trabajo,id_acceso,nombre_completo FROM view_reporte_asistencia WHERE ruc='" & KEY_RUC & "' and dni='" & Trim(Me.DtcTrabajadorReporte.BoundText) & "' and fecha>='" & Format(fecha_ini, "YYYY-mm-dd") & "'  and fecha<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY id"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_asistencia", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdsalir_Click()
Me.frmpermiso.Visible = False
End Sub

Private Sub DtcTrabajadorReporte_Change()
Call llenar_datos_sueldo(Me.DtcTrabajadorReporte.BoundText)
End Sub
Public Sub llenar_datos_sueldo(ByVal ndni As String)
 Dim sabados As Integer, domingos As Integer
Dim dias As Integer
Dim fecha As String
Dim horas_laborables As Single
strCadena = "SELECT sueldo FROM entidad_empresa WHERE cod_unico='" & ndni & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   
   
sabados = 0
domingos = 0
ndias = Val(DateSerial(Trim(Me.txtanio.Text), Me.dtcmes.BoundText + 1, 0))
For i = 1 To ndias
     fecha = Format(Format(i, "00") & "-" & Format(Me.dtcmes.BoundText, "00") & "-" & Trim(Me.txtanio.Text), "YYYY-mm-dd")
    If UCase(WeekdayName(Weekday(fecha))) = "SÁBADO" Then
       sabados = sabados + 1
    End If
    
    If UCase(WeekdayName(Weekday(fecha))) = "DOMINGO" Then
       domingos = domingos + 1
    End If
Next i



dias = (ndias - domingos)
horas_laborables = dias * 8

Me.txtprecio_hora.Text = Format(rst("sueldo") / horas_laborables, "###0.00")


End If
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpInicio.Value = KEY_FECHA

strCadena = "SELECT e.cod_unico as Codigo,p.nombre_completo as Descripcion FROM entidad_empresa e,persona p where e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTrabajador)

strCadena = "SELECT id_turno as Codigo,descripcion as Descripcion FROM turno_hc WHERE ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTurno)


strCadena = "SELECT id_mes Codigo,descripcion as Descripcion FROM meses"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcmes)
Me.dtcmes.BoundText = Format(Month(KEY_FECHA), "00")
Me.txtanio.Text = Year(KEY_FECHA)
strCadena = "SELECT e.cod_unico as Codigo,p.nombre_completo as Descripcion FROM entidad_empresa e,persona p where e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTrabajadorReporte)



strCadena = "SELECT * FROM view_permiso WHERE ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfdPersona)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
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
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 4000
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 2500
     
          Next
         cabecera = "CODIGO" & vbTab & "FECHA PERMISO" & vbTab & "TRABAJADO" & vbTab & "MOTIVO" & vbTab & "AUTORIZADO POR" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             
            Fila = rst("id_permiso") & vbTab & Format(rst("fecha_permiso"), "dd-mm-YYYY") & vbTab & rst("nombre_completo") & vbTab & rst("detalle") & vbTab & rst("autorizado") & vbTab & rst("usuario")
            Grilla.AddItem Fila
            Fila = ""
        rst.MoveNext
        Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbuscar_Change()
strCadena = "SELECT e.cod_unico as Codigo,p.nombre_completo as Descripcion FROM entidad_empresa e,persona p where p.nombre_completo LIKE '%" & Trim(Me.TXTBUSCAR.Text) & "%' and  e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTrabajador)

End Sub

Private Sub txtbuscar_trabajador_Change()
strCadena = "SELECT e.cod_unico as Codigo,p.nombre_completo as Descripcion FROM entidad_empresa e,persona p where p.nombre_completo LIKE '%" & Trim(Me.txtbuscar_trabajador.Text) & "%' and  e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTrabajadorReporte)

End Sub

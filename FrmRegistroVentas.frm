VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRegistroVentas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
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
   Begin VB.Frame frmResumen 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "AJUSTE TIPO CAMBIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2055
      Left            =   13200
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   5415
      Begin MSDataListLib.DataCombo DtcPeriodoResumen 
         Height          =   330
         Left            =   1320
         TabIndex        =   34
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin VitekeySoft.ChameleonBtn cmdResumenDetallado 
         Height          =   555
         Left            =   1320
         TabIndex        =   35
         Top             =   1320
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   979
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentas.frx":0000
         PICN            =   "FrmRegistroVentas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4920
         Picture         =   "FrmRegistroVentas.frx":25ED
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   2055
         Left            =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PERIODO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   360
         TabIndex        =   36
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.Frame frmajustebanco 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "AJUSTE TIPO CAMBIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2055
      Left            =   13200
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   1320
         TabIndex        =   30
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   435
         Left            =   1320
         TabIndex        =   31
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
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
         MICON           =   "FrmRegistroVentas.frx":5491
         PICN            =   "FrmRegistroVentas.frx":54AD
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
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PERIODO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   690
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808080&
         Height          =   2055
         Left            =   0
         Top             =   0
         Width           =   5415
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4920
         Picture         =   "FrmRegistroVentas.frx":7A92
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame frmImportacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTACION DE VENTAS"
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
      Height          =   6255
      Left            =   1560
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   16815
      Begin VitekeySoft.ChameleonBtn cmdImportarDesdeKeyfacil 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   5520
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "        IMPORTAR VENTAS        "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentas.frx":A936
         PICN            =   "FrmRegistroVentas.frx":A952
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   840
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51052545
         CurrentDate     =   43631
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "IMPORTACION PLANTILLA"
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
         Height          =   280
         Left            =   8760
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "IMPORTACION KEYFACIL"
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
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   51052545
         CurrentDate     =   43631
      End
      Begin MSComctlLib.ProgressBar prog_indicador 
         Height          =   220
         Left            =   120
         TabIndex        =   20
         Top             =   5230
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImportacionPlantilla 
         Height          =   4095
         Left            =   8760
         TabIndex        =   22
         Top             =   1080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7223
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   8760
         TabIndex        =   23
         Top             =   5235
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmdImportarPlantilla 
         Height          =   495
         Left            =   8760
         TabIndex        =   24
         Top             =   5550
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "        IMPORTAR VENTAS        "
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentas.frx":E693
         PICN            =   "FrmRegistroVentas.frx":E6AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn3 
         Height          =   300
         Left            =   8760
         TabIndex        =   25
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "CARGAR ARCHIVO"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentas.frx":123F0
         PICN            =   "FrmRegistroVentas.frx":1240C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImportacionkeyfacil 
         Height          =   4095
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7223
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
      Begin VitekeySoft.ChameleonBtn cmdProcesarPlantilla 
         Height          =   495
         Left            =   12000
         TabIndex        =   27
         Top             =   5550
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "PROCESAR VENTAS"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentas.frx":151A2
         PICN            =   "FrmRegistroVentas.frx":151BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image cmdclose 
         Height          =   240
         Left            =   16440
         Picture         =   "FrmRegistroVentas.frx":18EFF
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DESDE :"
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
         Left            =   180
         TabIndex        =   17
         Top             =   765
         Width           =   510
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   750
      Left            =   18840
      TabIndex        =   6
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":1BDA3
      PICN            =   "FrmRegistroVentas.frx":1BDBF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtRuc 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "txtRuc"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "TxtEmpresa"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSChart20Lib.MSChart Chart 
      Height          =   4500
      Left            =   120
      OleObjectBlob   =   "FrmRegistroVentas.frx":1C211
      TabIndex        =   5
      Top             =   4680
      Width           =   18615
   End
   Begin VitekeySoft.ChameleonBtn cmdIngresar 
      Height          =   750
      Left            =   18840
      TabIndex        =   7
      Top             =   1365
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "INGRESAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":1FAF9
      PICN            =   "FrmRegistroVentas.frx":1FB15
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   750
      Left            =   18840
      TabIndex        =   8
      Top             =   7680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":2214E
      PICN            =   "FrmRegistroVentas.frx":2216A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSunat 
      Height          =   750
      Left            =   18840
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "PLE SUNAT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":2255A
      PICN            =   "FrmRegistroVentas.frx":22576
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmExportar 
      Height          =   750
      Left            =   18840
      TabIndex        =   10
      Top             =   2895
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "EXPORTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":258C4
      PICN            =   "FrmRegistroVentas.frx":258E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmImportar 
      Height          =   750
      Left            =   18840
      TabIndex        =   11
      Top             =   2130
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "IMPORTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":29621
      PICN            =   "FrmRegistroVentas.frx":2963D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   750
      Left            =   18840
      TabIndex        =   12
      Top             =   6885
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "IMPRIMIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":2D37E
      PICN            =   "FrmRegistroVentas.frx":2D39A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdExportardia 
      Height          =   750
      Left            =   18840
      TabIndex        =   13
      Top             =   3660
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "CONSOLIDADO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":2F96B
      PICN            =   "FrmRegistroVentas.frx":2F987
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCostos 
      Height          =   750
      Left            =   18840
      TabIndex        =   28
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "COSTOS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":3227F
      PICN            =   "FrmRegistroVentas.frx":3229B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdResumen 
      Height          =   750
      Left            =   18840
      TabIndex        =   37
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "RESUMEN"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "FrmRegistroVentas.frx":34B85
      PICN            =   "FrmRegistroVentas.frx":34BA1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   7035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   750
      TabIndex        =   3
      Top             =   120
      Width           =   2085
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   240
      Top             =   60
      Width           =   18135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede


Public Sub actualizar_anio(ByVal ruc As String, ByVal Anio As String)
'strCadena = "SELECT ruc,mes,descripcion AS Periodo,anio, estado  as Estado FROM  dbo.RegistroVentas WHERE ruc='" & Trim(Me.txtruc.Text) & "' AND anio LIKE '%" & Trim(Me.txtanio.Text) & "%' ORDER BY anio,mes"
'Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub ChameleonBtn8_Click()

End Sub

Private Sub cmdClose_Click()
Me.frmImportacion.Visible = False
End Sub

Private Sub cmdCostos_Click()



strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  
 

  Me.frmajustebanco.Visible = True
End Sub

Private Sub cmdExportardia_Click()

frmUtilidad.Show
in_anio = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)
in_mes = Format(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1), "00")

Call frmUtilidad.load_periodo(in_anio, in_mes)


End Sub

Private Sub cmdImportarKeyfacil_Click()

' INICIAR LA MIGRACION
'Call get_importar_pedido(Me.DTPicker1.Value, DateAdd("d", 2, Me.DTPicker1.Value))


End Sub

Private Sub cmdImportarDesdeKeyfacil_Click()

Call get_importar_keyfacil(Me.DtpInicio.Value, DateAdd("d", 2, Me.DtpFin.Value))


End Sub




Private Sub cmdImprimir_Click()

strCadena = "SELECT periodo,anio,mes,acumulado FROM view_registro_ventas WHERE ruc='" & KEY_RUC & "'ORDER BY mes ASC "
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_registro_ventas", , App.Path + "\Reportes\")

End Sub

Private Sub cmdingresar_Click()
      Procedencia = modificar

      FrmRegistroVentasList.Show
End Sub

Private Sub cmdnuevo_Click()
      Procedencia = nuevo
      FrmDetalleRegistroVentas.Show
End Sub

Private Sub cmdProcesar_Click()

If MsgBox("Se va a procesar e Costo de Periodo:" & Space(1) & Me.DtcPeriodo.Text, vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then

    
    strCadena = "call CON_Costo_Internacional('1','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    MsgBox "Proceso Correcto..", vbInformation

End If


End Sub

Private Sub cmdResumen_Click()
strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoResumen)
   DtcPeriodoResumen.BoundText = get_periodo_actual(KEY_FECHA)
 
Me.frmResumen.Visible = True
End Sub

Private Sub cmdResumenDetallado_Click()
Dim cam3(0 To 2, 1 To 2)  As String



                    cam3(0, 1) = "in_fecha_ini"
                    cam3(1, 1) = "in_fecha_fin"
                    cam3(2, 1) = "in_vendedor"
                    
                   
    
                    cam3(0, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
                    cam3(1, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
                    cam3(2, 2) = Me.DtcPeriodoResumen.Text
                    param = cam3()
                  
            
                strCadena = "call CON_ResumenVenta('1','" & Me.DtcPeriodoResumen.BoundText & "','" & KEY_RUC & "')"
                
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptResumenVenta", param, App.Path + "\Reportes\")
                       
                       
                       




End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSunat_Click()
        
strCadena = "SELECT * FROM con_periodo WHERE mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' and Ejercicio='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   
   Call CreateXmlVentasCompras(rst("Id"))
   
End If






Exit Sub
        
        
        Procedencia = nuevo
        FrmRegistroSunat.Show
End Sub

Private Sub cmdTest_Click()
End Sub

Private Sub cmExportar_Click()
 Procedencia = nuevo
 frmTempExportExcel.Show
End Sub

Private Sub cmImportar_Click()
 
 Me.DtpInicio.Value = KEY_FECHA
 Me.DtpFin.Value = KEY_FECHA
 Me.HfImportacionkeyfacil.Rows = 0
 Me.HfImportacionPlantilla.Rows = 0
 Me.frmImportacion.Visible = True
 
 
 
 
 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.LblEmpresa.Caption = KEY_EMPRESA + Space(2) + "***[" + "RUC:" + Space(2) + KEY_RUC + "]***"
Me.TxtEmpresa.Text = KEY_EMPRESA
Me.txtRuc.Text = KEY_RUC
Call actualizar

End Sub


Public Sub actualizar()


strCadena = "SELECT * FROM view_registro_ventas WHERE ruc='" & KEY_RUC & "' "

Call llenarGrid(Me.HfdPersona, Me)




End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
Public Sub ImportarVentas(ByVal ruc As String)
Dim rstRemoto As New ADODB.Record
  Set rstT = New ADODB.Recordset
  rstT.CursorLocation = adUseClient
  strCadena = "SELECT * FROM RegistroVentasDetalle WHERE mes='01'OR mes='02' AND anio='2013' WHERE ruc='20104050337' ORDER BY  codigounico ASC"
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
        strCadena = "P_insert_venta('" & formato_item(rstT("doc_cod"), 4) & "','00001','" & formato_item(rstT("idformapago"), 2) & "','" & rstT("moneda") & "','no'," & _
            "'" & formato_item(Trim(rstT("serie")), 3) & "','" & formato_item(Trim(rstT("numero")), 6) & "','" & rstT("RucCliente") & "','" & rstT("NombreCliente") & "','" & rstT("afecto") & "','" & rstT("igv") & "','" & rstT("exonerado") & "','" & rstT("total") & "','0'," & _
            "'" & rstT("total") & "','0','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & rstT("tc") & "','no','" & rstT("mes") & "','" & rstT("anio") & "','" & ruc & "')"
            CnBd.Execute (strCadena)
             
            
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            If rstT("anulado") = "V" Then
                strCadena = "UPDATE movimiento_venta SET anulado='si',ncliente='A N U L A D O',id_cliente='',afecto='0',exonerado='0',igv='0',total='0',saldo='0' WHERE id_venta='" & id_venta & "'"
                CnBd.Execute (strCadena)
                 
            End If
            rstT.MoveNext
            DoEvents
    Next i
End If
End Sub

Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim Acumulado As Double
 Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
  
  
    Grilla.Rows = 0
    Chart.Visible = True
    Chart.TitleText = "VARIACION DE VENTAS"
    Me.Chart.RowCount = rst.RecordCount
      
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1800
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 2000
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 2000
        Next
      
        cabecera = "RUC" & vbTab & "MES" & vbTab & "PERIODO" & vbTab & "AÑO" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "TOTAL VENTA" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 7
             Grilla.col = k
             Grilla.Row = 0
             Grilla.CellBackColor = &HDFDFE0
        Next k

        rst.MoveFirst
        
        in_valor_ventat = 0
        in_igvt = 0
        in_totalt = 0
        For i = 0 To rst.RecordCount - 1
        
        If KEY_PAIS <> KEY_PERU Then
            If IsNull(rst("acumulado_inter")) = True Then
                in_acumulado = 0
            Else
                in_acumulado = rst("acumulado_inter")
            End If
        Else
            If IsNull(rst("acumulado")) = True Then
                in_acumulado = 0
            Else
                in_acumulado = rst("acumulado")
            End If
        End If
        
        
        If KEY_CON_IGV = "si" Then
           in_valor_venta = in_acumulado / (1 + KEY_IGV)
           in_igv = in_valor_venta * KEY_IGV
        Else
           in_valor_venta = in_acumulado
           in_igv = 0
        End If
        in_valor_ventat = in_valor_ventat + in_valor_venta
        in_igvt = in_igv + in_igvt
        in_totalt = in_totalt + in_acumulado
            Fila = rst("ruc") & vbTab & rst("mes") & vbTab & rst("periodo") & vbTab & rst("anio") & vbTab & Format(in_valor_venta, "#,##0.000") & vbTab & Format(in_igv, "#,##0.000") & vbTab & Format(in_acumulado, "#,##0.000") & vbTab & rst("estado")
            Grilla.AddItem Fila
               
                    For j = 4 To 7
                       Grilla.col = j
                       Grilla.Row = i + 1
                       If (rst("estado") = "PENDIENTE") Then
                            Grilla.CellBackColor = &H8080FF
                        Else
                            Grilla.CellBackColor = &HC0FFC0
                        End If
                    Next j
                    
                    
                   
    
        ' Establecemos las Etiquetas de las Columnas
        Chart.DataGrid.RowLabel(i + 1, 1) = Mid(rst("nmes"), 1, 5)
        Chart.DataGrid.SetSize rst.RecordCount, 1, rst.RecordCount, 1
        Chart.DataGrid.SetData i + 1, 1, in_acumulado, 0
    
    
    
               
        
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_valor_ventat, "#,##0.000") & vbTab & Format(in_igvt, "#,##0.000") & vbTab & Format(in_totalt, "#,##0.000")
        Grilla.AddItem Fila
  
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub txtAnio_Change()
'If (Len(Me.txtanio.Text) = 4) Then
'    Me.Command1.Enabled = True
'Else
 '   Me.Command1.Enabled = False
'End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
 '   Call actualizar_anio(Me.txtRuc.Text, Me.txtanio.Text)
'End If
End Sub




Private Sub Image1_Click()
Me.frmajustebanco.Visible = False
End Sub

Private Sub Image2_Click()
Me.frmResumen.Visible = False
End Sub

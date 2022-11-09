VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmHistoriaClinica 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   19845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtAnio 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4080
      MaxLength       =   11
      TabIndex        =   78
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1710
      MaxLength       =   50
      TabIndex        =   33
      Top             =   840
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3855
      Left            =   10080
      TabIndex        =   30
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "ANT. FAMILIARES"
      TabPicture(0)   =   "FrmHistoriaClinica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "HfgAntecedentesFam"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ANT. PATOLOGICOS"
      TabPicture(1)   =   "FrmHistoriaClinica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "HfgAntecedentesPatolo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ANT.QUIRURGICO"
      TabPicture(2)   =   "FrmHistoriaClinica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "HfgAntecedentesQuirur"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "VACUNAS APLICADAS"
      TabPicture(3)   =   "FrmHistoriaClinica.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "HgfVacunas"
      Tab(3).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAntecedentesFam 
         Height          =   2535
         Left            =   240
         TabIndex        =   71
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAntecedentesPatolo 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   72
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAntecedentesQuirur 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   73
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HgfVacunas 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   74
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   9960
      TabIndex        =   16
      Top             =   4200
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   8
      Tab             =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "DIRECCION"
      TabPicture(0)   =   "FrmHistoriaClinica.frx":0070
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TxtDireccion1"
      Tab(0).Control(1)=   "TxtDistrito"
      Tab(0).Control(2)=   "DtcDistrito"
      Tab(0).Control(3)=   "DtcUrbanizacion"
      Tab(0).Control(4)=   "DtcZona"
      Tab(0).Control(5)=   "DtcDepartamento"
      Tab(0).Control(6)=   "DtcProvincia"
      Tab(0).Control(7)=   "LblDireccion"
      Tab(0).Control(8)=   "lbldistrito"
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "lbldepartamento"
      Tab(0).Control(12)=   "lblprovincia"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "T.SANGRE/ ALERGIAS"
      TabPicture(1)   =   "FrmHistoriaClinica.frx":008C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DtcGruposangre"
      Tab(1).Control(1)=   "DtcFactorRh"
      Tab(1).Control(2)=   "SSTab4"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Shape1"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "IMC"
      TabPicture(2)   =   "FrmHistoriaClinica.frx":00A8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape2"
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(3)=   "SSTab5"
      Tab(2).Control(4)=   "TxtEstatura"
      Tab(2).Control(5)=   "TxtPeso"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "PRESION ARTERIAL"
      TabPicture(3)   =   "FrmHistoriaClinica.frx":00C4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "hfgPresionarterial"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "FAMILIARES"
      TabPicture(4)   =   "FrmHistoriaClinica.frx":00E0
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Shape3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label11"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label14"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label15"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label16"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label17"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label18"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "DtcParentesco"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "HfgFamiliares"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtDni"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "TxtFmaterno"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cmdAgregar"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "TxtTelefono"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "TxtFpaterno"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "TxtFnombers"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).ControlCount=   15
      TabCaption(5)   =   "ENFERMED CRONICAS"
      TabPicture(5)   =   "FrmHistoriaClinica.frx":00FC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "HfgEnfermedad"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "SEGURO MEDICO"
      TabPicture(6)   =   "FrmHistoriaClinica.frx":0118
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "HfgSeguro"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "GLUCOSA"
      TabPicture(7)   =   "FrmHistoriaClinica.frx":0134
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "HfgGlucosa"
      Tab(7).ControlCount=   1
      Begin VB.TextBox TxtFnombers 
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
         Left            =   1080
         TabIndex        =   69
         Top             =   4420
         Width           =   1815
      End
      Begin VB.TextBox TxtFpaterno 
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
         Left            =   1080
         TabIndex        =   68
         Top             =   4125
         Width           =   1815
      End
      Begin VB.TextBox TxtTelefono 
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
         Left            =   5160
         TabIndex        =   64
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   62
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox TxtFmaterno 
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
         Left            =   1080
         TabIndex        =   61
         Top             =   3820
         Width           =   1815
      End
      Begin VB.TextBox txtDni 
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
         Left            =   1080
         TabIndex        =   58
         Top             =   3520
         Width           =   1815
      End
      Begin VB.TextBox TxtPeso 
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
         Left            =   -69720
         MaxLength       =   11
         TabIndex        =   45
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox TxtEstatura 
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
         Left            =   -73320
         MaxLength       =   11
         TabIndex        =   44
         Top             =   800
         Width           =   855
      End
      Begin VB.TextBox TxtDireccion1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   -73350
         MaxLength       =   150
         TabIndex        =   28
         Top             =   2760
         Width           =   4935
      End
      Begin VB.TextBox TxtDistrito 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -70530
         MaxLength       =   80
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcDistrito 
         Height          =   315
         Left            =   -73350
         TabIndex        =   18
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcUrbanizacion 
         Height          =   315
         Left            =   -73350
         TabIndex        =   19
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcZona 
         Height          =   315
         Left            =   -73350
         TabIndex        =   20
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcDepartamento 
         Height          =   315
         Left            =   -73350
         TabIndex        =   21
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcProvincia 
         Height          =   315
         Left            =   -73350
         TabIndex        =   22
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcGruposangre 
         Height          =   315
         Left            =   -73080
         TabIndex        =   34
         Top             =   800
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSDataListLib.DataCombo DtcFactorRh 
         Height          =   315
         Left            =   -69315
         TabIndex        =   36
         Top             =   800
         Width           =   1335
         _ExtentX        =   2355
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
      Begin TabDlg.SSTab SSTab4 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   38
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5741
         _Version        =   393216
         TabOrientation  =   2
         Tab             =   2
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
         TabCaption(0)   =   "Medicame"
         TabPicture(0)   =   "FrmHistoriaClinica.frx":0150
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "HfdMedicamento"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Alimentos"
         TabPicture(1)   =   "FrmHistoriaClinica.frx":016C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "HfgAlimento"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Ambiente"
         TabPicture(2)   =   "FrmHistoriaClinica.frx":0188
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "HfgAmbiente"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdMedicamento 
            Height          =   2295
            Left            =   -74520
            TabIndex        =   39
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAlimento 
            Height          =   2295
            Left            =   -74520
            TabIndex        =   40
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAmbiente 
            Height          =   2295
            Left            =   480
            TabIndex        =   41
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
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
      End
      Begin TabDlg.SSTab SSTab5 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   46
         Top             =   1320
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5741
         _Version        =   393216
         TabOrientation  =   2
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
         TabCaption(0)   =   "HISTORIAL"
         TabPicture(0)   =   "FrmHistoriaClinica.frx":01A4
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Chart"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "IMC"
         TabPicture(1)   =   "FrmHistoriaClinica.frx":01C0
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label10"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lblimc"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "HfgImc"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2295
            Left            =   -74520
            TabIndex        =   47
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   2295
            Left            =   -74520
            TabIndex        =   48
            Top             =   360
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4048
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgImc 
            Height          =   2655
            Left            =   480
            TabIndex        =   49
            Top             =   480
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   4683
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
         Begin MSChart20Lib.MSChart Chart 
            Height          =   2895
            Left            =   -74520
            OleObjectBlob   =   "FrmHistoriaClinica.frx":01DC
            TabIndex        =   52
            Top             =   120
            Visible         =   0   'False
            Width           =   7335
         End
         Begin VB.Label lblimc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2760
            TabIndex        =   51
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INDICE DE MASA MUSCULAR :"
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
            Left            =   435
            TabIndex        =   50
            Top             =   120
            Width           =   2205
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgPresionarterial 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   53
         Top             =   960
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgGlucosa 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   54
         Top             =   960
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSeguro 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   55
         Top             =   840
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5318
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFamiliares 
         Height          =   2535
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
      Begin MSDataListLib.DataCombo DtcParentesco 
         Height          =   315
         Left            =   5160
         TabIndex        =   60
         Top             =   3720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgEnfermedad 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   70
         Top             =   840
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4471
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
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES :"
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
         TabIndex        =   67
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.PATERNO :"
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
         Left            =   165
         TabIndex        =   66
         Top             =   4125
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.MATERNO :"
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
         Left            =   135
         TabIndex        =   65
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO :"
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
         Left            =   4080
         TabIndex        =   63
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENTESCO :"
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
         Left            =   3870
         TabIndex        =   59
         Top             =   3840
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI :"
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
         TabIndex        =   57
         Top             =   3600
         Width           =   405
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   1270
         Left            =   60
         Top             =   3480
         Width           =   8055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTATURA (m) :"
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
         Left            =   -74625
         TabIndex        =   43
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESO (KG) :"
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
         Left            =   -70710
         TabIndex        =   42
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTOR RH :"
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
         Left            =   -70770
         TabIndex        =   37
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRUPO SANGUINEO :"
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
         Left            =   -74820
         TabIndex        =   35
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblDireccion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección 1 :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   29
         Top             =   2835
         Width           =   915
      End
      Begin VB.Label lbldistrito 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRITO :"
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
         Left            =   -74325
         TabIndex        =   27
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "URBANIZACION :"
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
         Left            =   -74760
         TabIndex        =   26
         Top             =   2085
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZONA :"
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
         Left            =   -73980
         TabIndex        =   25
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label lbldepartamento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTAMENTO :"
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
         Left            =   -74820
         TabIndex        =   24
         Top             =   1005
         Width           =   1335
      End
      Begin VB.Label lblprovincia 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA :"
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
         Left            =   -74430
         TabIndex        =   23
         Top             =   1365
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   -74880
         Top             =   720
         Width           =   8055
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   -74880
         Top             =   720
         Width           =   8055
      End
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1710
      MaxLength       =   80
      TabIndex        =   6
      Top             =   3240
      Width           =   4215
   End
   Begin VB.CommandButton CmdFoto 
      Caption         =   "CAPTURAR IMAGEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6030
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtdia 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1710
      MaxLength       =   11
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtPaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1710
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1710
      MaxLength       =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtMaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1710
      MaxLength       =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   4215
   End
   Begin VitekeySoft.TextBoxPlus txtRazonSocial 
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   11760
      Top             =   5760
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
            Picture         =   "FrmHistoriaClinica.frx":3AC4
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":3DE0
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":4240
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":46A0
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":49BC
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":4E1C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":5138
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":5598
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":59F8
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":62D8
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":65F4
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinica.frx":6910
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -1050
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo Dtcmes 
      Height          =   315
      Left            =   2430
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin InetCtlsObjects.Inet inetConecta 
      Left            =   -570
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   4935
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8705
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "CONSULTAS"
      TabPicture(0)   =   "FrmHistoriaClinica.frx":6C2C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "HfgConsultas"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "EXAMENES CLINICOS"
      TabPicture(1)   =   "FrmHistoriaClinica.frx":6C48
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "HfgExamenes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "RECETAS"
      TabPicture(2)   =   "FrmHistoriaClinica.frx":6C64
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "HfgRecetas"
      Tab(2).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgConsultas 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   75
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7435
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgExamenes 
         Height          =   4215
         Left            =   120
         TabIndex        =   76
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   7435
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgRecetas 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   77
         Top             =   480
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7646
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
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
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
      Left            =   1320
      TabIndex        =   32
      Top             =   840
      Width           =   405
   End
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE COMPLETO :"
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
      TabIndex        =   15
      Top             =   2340
      Width           =   1605
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL :"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   3300
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   3540
      Left            =   6150
      Picture         =   "FrmHistoriaClinica.frx":6C80
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblCumpleaños 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA NACIMIENTO :"
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
      Left            =   90
      TabIndex        =   13
      Top             =   2805
      Width           =   1635
   End
   Begin VB.Label LblCodPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   1710
      TabIndex        =   12
      Top             =   75
      Width           =   3375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD.BARRA :"
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
      Left            =   690
      TabIndex        =   11
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRES :"
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
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A.PATERNO :"
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
      Left            =   720
      TabIndex        =   9
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A.MATERNO :"
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
      Left            =   690
      TabIndex        =   8
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4035
      Left            =   30
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "FrmHistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub CmdAgregar_Click()
Dim razon As String
If Me.txtDni.Text <> "" Then
    strCadena = "SELECT * FROM persona_accidentes WHERE dni_familia='" & Trim(Me.txtDni.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "UPDATE persona_accidentes SET telefono='" & Me.TxtTelefono.Text & "',id_parentesco='" & Me.DtcParentesco.BoundText & "' WHERE dni_familia='" & Me.txtDni.Text & "' AND dni='" & Me.txtruc.Text & "'"
        CnBd.Execute (strCadena)
        Call insertar_acciones(strCadena)
    Else
        strCadena = "INSERT INTO persona_accidentes(dni,dni_familia,id_parentesco,telefono)VALUES('" & Me.txtruc.Text & "','" & Me.txtDni.Text & "','" & Me.DtcParentesco.BoundText & "','" & Me.TxtTelefono.Text & "')"
        CnBd.Execute (strCadena)
        Call insertar_acciones(strCadena)
        strCadena = "select * from persona where dni='" & Trim(Me.txtDni.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
        razon = Me.TxtFnombers.Text + Space(1) + Me.TxtFpaterno.Text + Space(1) + Me.TxtFmaterno.Text
        strCadena = "P_insert_persona('" & Trim(Me.txtDni.Text) & "','" & Me.TxtFpaterno.Text & "','" & Me.TxtFmaterno.Text & "','" & Me.TxtFnombers.Text & "','" & Trim(razon) & "','CHICLAYO','" & Me.TxtTelefono.Text & "','--','no','no','no','no','no','0','')"
        CnBd.Execute (strCadena)
        Call insertar_acciones(strCadena)
        End If
    End If
    strCadena = "SELECT F.id,P.nombre_completo,PR.descripcion as parentesco,F.telefono,F.dni_familia FROM persona_accidentes F,persona P,parentesco PR WHERE F.id_parentesco=PR.id_parentesco AND  F.dni='" & Me.txtruc.Text & "' AND F.dni_familia=P.dni"
    Call llenarFamiliares(Me.HfgFamiliares)
End If
End Sub



Private Sub DtcDepartamento_Change()
strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcProvincia)

End Sub

Private Sub DtcDistrito_Change()
strCadena = "SELECT id_urbanizacion as Codigo,descripcion as Descripcion FROM urbanizacion WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcUrbanizacion)
End Sub

Private Sub DtcParentesco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtTelefono)
End If
End Sub

Private Sub DtcProvincia_Change()
strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcDistrito)
End Sub

Private Sub DtcUrbanizacion_Change()
If Me.DtcUrbanizacion.BoundText <> "" Then
strCadena = "SELECT  id_zona as Codigo,descripcion_zona as Descripcion FROM zona WHERE id_Urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcZona)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub


Private Sub Form_Load()
CenterForm Me
Me.Top = 50
strCadena = "SELECT id_mes as Codigo,descripcion as Descripcion FROM meses ORDER BY id_mes"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)

strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDepartamento)
strCadena = "SELECT id_grupo as Codigo,descripcion as Descripcion FROM hc_grupo_sanguineo ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcGruposangre)

strCadena = "SELECT id_factor as Codigo,descripcion as Descripcion FROM hc_factorrh ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFactorRh)

strCadena = "SELECT id_parentesco as Codigo,descripcion as Descripcion FROM parentesco ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcParentesco)
If FrmPersona.Procedencia = Selecionar Then
    Call llenar_datos(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0))
End If
    

End Sub
Public Sub llenar_datos(ByVal dni As String)
    strCadena = "SELECT * FROM persona WHERE dni='" & dni & "'"
    Call ConfiguraRst(strCadena)
    Me.LblCodPersona.Caption = rst("dni")
    Me.txtruc.Text = rst("dni")
    Me.txtPaterno.Text = UCase(rst("a_paterno"))
    Me.txtMaterno.Text = UCase(rst("a_materno"))
    Me.txtNombre.Text = UCase(rst("nombres"))
    Me.txtrazonsocial.Text = UCase(rst("nombre_completo"))
    Me.txtdia.Text = rst("id_dia")
    Me.DtcMes.BoundText = formato_item(rst("id_mes"), 2)
    Me.txtAnio.Text = rst("id_anio")
    Me.TxtEmail.Text = rst("mail")
    Me.DtcDepartamento.BoundText = rst("id_departamento")
    Me.DtcProvincia.BoundText = rst("id_provincia")
    Me.DtcDistrito.BoundText = rst("id_distrito")
    Me.DtcUrbanizacion.BoundText = rst("id_urbanizacion")
    Me.DtcZona.BoundText = rst("id_zona")
    Me.TxtDireccion1.Text = UCase(rst("direccion"))
    Me.DtcFactorRh.BoundText = rst("id_factor")
    Me.DtcGruposangre.BoundText = rst("id_grupo_sangre")
    Me.TxtEstatura.Text = Format(rst("estatura"), "#,##0.00")
    Me.TxtPeso.Text = Format(rst("peso"), "#,##0.00")
    If rst("estatura") > 0 Then
    Me.lblimc.Caption = Format(rst("peso") / (rst("estatura") * rst("estatura")), "#,##0.00")
    End If
    '--------- foto--------
If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
    If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
        Me.Image1.Visible = True
        Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
        img = Trim(rst("foto"))
    Else
        Me.Image1.Visible = False
    End If
    
    strCadena = "SELECT id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_alergia_medi M,persona P WHERE M.dni='" & dni & "' AND M.dni_save=P.dni"
    Call llenarAlergias(Me.HfdMedicamento)
    
    strCadena = "SELECT id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_aler_alimento M,persona P WHERE M.dni='" & dni & "' AND M.dni_save=P.dni"
    Call llenarAlergias(Me.HfgAlimento)
    
    strCadena = "SELECT id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_aler_ambiente M,persona P WHERE M.dni='" & dni & "' AND M.dni_save=P.dni"
    Call llenarAlergias(Me.HfgAmbiente)
    Call llenarImc(Me.HfgImc)
    Call llenar_imc_historial(dni)
    
    strCadena = "SELECT id as id_codigo,fecha,P.nombre_completo as realizado,sistolica,diastolica FROM persona_presion_art M,persona P WHERE M.dni='" & dni & "' AND M.dni_save=P.dni ORDER BY M.fecha"
    Call llenarPresion(Me.hfgPresionarterial)
    
    strCadena = "SELECT id as id_codigo,fecha,P.nombre_completo as realizado,valor as descripcion FROM persona_glucosa M,persona P WHERE M.dni='" & dni & "' AND M.dni_save=P.dni ORDER BY M.fecha"
    Call llenarAlergias(Me.HfgGlucosa)
    strCadena = "SELECT * FROM persona_seguro WHERE dni='" & dni & "'"
    Call llenarSeguro(Me.HfgSeguro)
    
    strCadena = "SELECT F.id,P.nombre_completo,PR.descripcion as parentesco,F.telefono,F.dni_familia FROM persona_accidentes F,persona P,parentesco PR WHERE F.id_parentesco=PR.id_parentesco AND  F.dni='" & dni & "' AND F.dni_familia=P.dni"
    Call llenarFamiliares(Me.HfgFamiliares)
    
    strCadena = "SELECT E.id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_enfermedad_cronica E,persona P WHERE E.dni='" & dni & "' AND E.dni_save=P.dni"
    Call llenarAlergias(Me.HfgEnfermedad)
    
    strCadena = "SELECT E.id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_antecedentes_fam E,persona P WHERE E.dni='" & dni & "' AND E.dni_save=P.dni"
    Call llenarAlergias(Me.HfgAntecedentesFam)
    
    strCadena = "SELECT E.id as id_codigo,descripcion,fecha,P.nombre_completo as realizado FROM persona_antecedente_patologico E,persona P WHERE E.dni='" & dni & "' AND E.dni_save=P.dni"
    Call llenarAlergias(Me.HfgAntecedentesPatolo)
    strCadena = "SELECT E.id as id_codigo,clinica,cirujano,cirugia,detalle,fecha,P.nombre_completo as realizado FROM persona_antecedente_quirurgico E,persona P WHERE E.dni='" & dni & "' AND E.dni_save=P.dni"
    Call llenarQuirurgico(Me.HfgAntecedentesQuirur)
    
    strCadena = "SELECT E.id as id_codigo,vacuna,fecha_reg,comentario,P.nombre_completo as realizado FROM persona_vacunas E,persona P WHERE E.dni='" & dni & "' AND E.dni_save=P.dni"
    Call llenarVacunas(Me.HgfVacunas)
    
    strCadena = "SELECT * FROM persona_consultas C,persona P WHERE C.dni='" & dni & "' AND C.dni_save=P.dni"
    Call llenarConsultas(Me.HfgConsultas)
    
    strCadena = "SELECT P.id_analisis,A.descripcion as Clasificacion,L.descripcion,P.dni_save,P.estado,P.fecha,P.fecha_atencion as fecha_resultado,P.ruc_empresa FROM persona_analisis P,analisis_clinico_listado L,analisis_clinico A WHERE P.id_analisis=L.id_analisis AND P.dni='" & dni & "' AND L.id_clasificacion=A.id_clasificacion ORDER BY A.id_clasificacion"
    Call llenarExamenes(Me.HfgExamenes)
End If
End Sub
Private Sub llenarExamenes(ByVal Grilla As MSHFlexGrid)
Dim color As String
Dim fecha_atencion  As String, Laboratorio As String

Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 3000
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 2500
           Grilla.ColWidth(6) = 2500
        Next
        cabecera = "ANALISIS" & vbTab & "FECHA" & vbTab & "CLASIFICACION" & vbTab & "DESCRIPCION" & vbTab & "FECHA RESULTADO" & vbTab & "LABORATORIO" & vbTab & "REQUERIDO POR"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          If IsNull(rstT("fecha_resultado")) = True Then
             fecha_atencion = ""
           Else
             fecha_atencion = rstT("fecha_resultado")
           End If
           
           If rstT("ruc_empresa") = "00000000" Then
              Laboratorio = ""
           Else
             Laboratorio = UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rstT("ruc_empresa")))
           End If
          Fila = rstT("id_analisis") & vbTab & rstT("fecha") & vbTab & UCase(rstT("Clasificacion")) & vbTab & rstT("descripcion") & vbTab & rstT("fecha_resultado") & vbTab & Laboratorio & vbTab & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rstT("dni_save")))
          Grilla.AddItem Fila
          If rstT("estado") = "Pendiente" Then
             color = &H8080FF
          Else
            color = &HC0FFC0
          End If
            For k = 0 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = color
         Next k
          
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenarConsultas(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 3700
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 3000
        
        Next
        cabecera = "NºCONSULTA" & vbTab & "MOTIVO" & vbTab & "FECHA" & vbTab & "DR.TRATANTE"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_consulta") & vbTab & UCase(rstT("motivo")) & vbTab & rstT("fecha") & vbTab & UCase(rstT("nombre_completo"))
          Grilla.AddItem Fila
          If rstT("estado") = "01" Then
            For k = 0 To 3
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
         Next k
          End If
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenarVacunas(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 2200
        Next
        cabecera = "IDCODIGO" & vbTab & "VACUNA" & vbTab & "DETALLE" & vbTab & "FECHA" & vbTab & "REGISTRADO POR"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_codigo") & vbTab & rstT("vacuna") & vbTab & UCase(rstT("comentario")) & vbTab & rstT("fecha_reg") & vbTab & UCase(rstT("realizado"))
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenarQuirurgico(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 2000
        Next
        cabecera = "IDCODIGO" & vbTab & "OPERACION" & vbTab & "DOCTOR" & vbTab & "CLINICA" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_codigo") & vbTab & rstT("fecha") & vbTab & UCase(rstT("cirujano")) & vbTab & UCase(rstT("clinica")) & vbTab & UCase(rstT("cirugia"))
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub
Private Sub llenarFamiliares(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1100
        Next
        cabecera = "IDCODIGO" & vbTab & "DNI" & vbTab & "FAMILIAR" & vbTab & "PARENTESCO" & vbTab & "TELEFONO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id") & vbTab & rstT("dni_familia") & vbTab & rstT("nombre_completo") & vbTab & rstT("parentesco") & vbTab & rstT("telefono")
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenarSeguro(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2800
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1100
        Next
        cabecera = "IDCODIGO" & vbTab & "ASEGURADORA" & vbTab & "Nº POLIZA" & vbTab & "EXPEDICION" & vbTab & "EXPIRACION"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id") & vbTab & rstT("descripcion") & vbTab & rstT("numero") & vbTab & rstT("expedicion") & vbTab & rstT("expiracion")
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenarPresion(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 4000
        Next
        cabecera = "IDCODIGO" & vbTab & "FECHA" & vbTab & "MEDICION" & vbTab & "REALZIADO POR"
        Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_codigo") & vbTab & rstT("fecha") & vbTab & rstT("sistolica") & "/" & rstT("diastolica") & Space(2) & "mm.Hg" & vbTab & rstT("realizado")
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenar_imc_historial(ByVal dni As String)
'Height = 5640: Width = 7575
Chart.Visible = True

'If Len(cboYear.text) = 4 Then
    ' Filtramos los Registros de la BD en base a el año seleccionado
    strCadena = "SELECT fecha,peso FROM persona_peso WHERE dni='" & dni & "'"
    Call ConfiguraRstT(strCadena)
    ' Calculamos el Porcentaje de las Ventas
    ' Esperadas y Logradas
    'nEsperado = (rsVentas.Fields("Esperado") / rsVentas.Fields("Esperado")) * 100
    'nLogrado = (rsVentas.Fields("Logrado") / rsVentas.Fields("Esperado")) * 100
    ' Colocamos el titulo al Chart
    Chart.TitleText = "VARIACION DE PESO"
    Me.Chart.RowCount = rstT.RecordCount
    With Chart.DataGrid
        ' Establecemos las Etiquetas de las Columnas
        For i = 1 To rstT.RecordCount
        .RowLabel(i, 1) = rstT("fecha")
        
        ' Establecemos el tamaño del Chart
        ' Parametros: Total Etiquetas Cols, Total Etiquetas Series,
        ' Total Columnas, Total Series
        .SetSize rstT.RecordCount, 1, rstT.RecordCount, 1
        ' Establecemos los valores de cada columna
        ' Parametros: Columna, Serie, Valor
        '.SetData 1, 1, nLogrado, 0
        '.SetData 2, 1, nEsperado, 0
        rstT.MoveNext
    Next i
    End With
'End If
End Sub

Private Sub llenarAlergias(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 3000
           Grilla.ColWidth(3) = 1100
        Next
        cabecera = "IDCODIGO" & vbTab & "DESCRIPCION" & vbTab & "AGREGADO POR" & vbTab & "FECHA"
        Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_codigo") & vbTab & rstT("descripcion") & vbTab & rstT("realizado") & vbTab & rstT("fecha")
          Grilla.AddItem Fila
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub
Private Sub llenarImc(ByVal Grilla As MSHFlexGrid)
Dim campos(1 To 4) As Integer, imc As Single
       Grilla.Clear
       Grilla.Refresh
       Grilla.Rows = 0
       ReDim arrColWidth(1 To 4)
       For Each Campo In campos
           Grilla.ColWidth(0) = 1500
           Grilla.ColWidth(1) = 2800
           Grilla.ColWidth(2) = 2800
        Next
        cabecera = "IMC" & vbTab & "CLASIFICACION" & vbTab & "RANGOS"
        Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
          Fila = "< 18.50" & vbTab & "PESO INSUFICIENTE" & vbTab & "Menos de 54.7 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) < 18.5 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 1
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          
          Fila = "18.5 - 24.99" & vbTab & "Peso Normal." & vbTab & "Entre 53.46 y 72.22 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 18.5 And Val(Me.lblimc.Caption) <= 24.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 2
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          
          Fila = "25.00 - 26.99" & vbTab & "Sobrepeso grado I " & vbTab & "Entre 72.25 y 78 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 25 And Val(Me.lblimc.Caption) <= 26.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 3
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          
          Fila = "27.00 - 29.99" & vbTab & "Sobrepeso grado II (preobesidad) " & vbTab & "Entre 78.03 y 86.67 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 27 And Val(Me.lblimc.Caption) <= 29.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 4
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          Fila = "30.00 - 34.99" & vbTab & "Obesidad de tipo I " & vbTab & "Entre 86.7 y 101.12 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 30 And Val(Me.lblimc.Caption) <= 34.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 5
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          Fila = "35.00 - 39.99" & vbTab & "Obesidad de tipo II " & vbTab & "Entre 101.15 y 115.57 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 35 And Val(Me.lblimc.Caption) <= 39.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 6
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          Fila = "40.00 - 49.99" & vbTab & "Obesidad de tipo III (mórbida) " & vbTab & "Entre 115.6 y 144.47 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) >= 40 And Val(Me.lblimc.Caption) <= 49.99 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 7
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          Fila = "> 50.00" & vbTab & "Obesidad de tipo IV (extrema) " & vbTab & "Más de 144.5 kg. "
          Grilla.AddItem Fila
          If Val(Me.lblimc.Caption) > 50 Then
              For k = 0 To 2
                  Grilla.col = k
                  Grilla.Row = 8
                  Grilla.CellBackColor = &H8080FF
             Next k
          End If
          
      
   
End Sub



Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_CANCEL
        Unload Me
End Select
End Sub

Private Sub HfgConsultas_DblClick()
If Val(Me.HfgConsultas.TextMatrix(Me.HfgConsultas.Row, 0)) > 0 Then
    Procedencia = Selecionar
    FrmHistoriaClinicaConsulta.Show
End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtFmaterno.Text = rst("a_materno")
       Me.TxtFpaterno.Text = rst("a_paterno")
       Me.TxtFnombers.Text = rst("nombres")
       Me.DtcParentesco.SetFocus
    Else
        Call Resalta(Me.TxtFpaterno)
    End If
    
End If
End Sub

Private Sub TxtFamiliar_Change()

End Sub

Private Sub TxtFamiliar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcParentesco.SetFocus
End If
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar.SetFocus
End If
End Sub

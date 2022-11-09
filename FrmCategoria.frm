VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmCategoria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TXTIDentificador 
      Height          =   375
      Left            =   9960
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   16351
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   10
      Tab             =   9
      TabsPerRow      =   10
      TabHeight       =   2117
      TabCaption(0)   =   "Nivel 9"
      TabPicture(0)   =   "FrmCategoria.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape19"
      Tab(0).Control(1)=   "Shape18"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(4)=   "HfNivel9"
      Tab(0).Control(5)=   "Text9"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Nivel 8"
      TabPicture(1)   =   "FrmCategoria.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape17"
      Tab(1).Control(1)=   "Shape16"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "HfNivel8"
      Tab(1).Control(5)=   "Text8"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Nivel 7"
      TabPicture(2)   =   "FrmCategoria.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape15"
      Tab(2).Control(1)=   "Shape14"
      Tab(2).Control(2)=   "Label13"
      Tab(2).Control(3)=   "Label14"
      Tab(2).Control(4)=   "HfNivel7"
      Tab(2).Control(5)=   "Text7"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Nivel 6"
      TabPicture(3)   =   "FrmCategoria.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Shape13"
      Tab(3).Control(1)=   "Shape12"
      Tab(3).Control(2)=   "Label11"
      Tab(3).Control(3)=   "Label12"
      Tab(3).Control(4)=   "HfNivel6"
      Tab(3).Control(5)=   "Text6"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Nivel 5"
      TabPicture(4)   =   "FrmCategoria.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Shape11"
      Tab(4).Control(1)=   "Shape10"
      Tab(4).Control(2)=   "Label9"
      Tab(4).Control(3)=   "Label10"
      Tab(4).Control(4)=   "HfNivel5"
      Tab(4).Control(5)=   "Text5"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Nivel 4"
      TabPicture(5)   =   "FrmCategoria.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Shape9"
      Tab(5).Control(1)=   "Shape8"
      Tab(5).Control(2)=   "Label7"
      Tab(5).Control(3)=   "Label8"
      Tab(5).Control(4)=   "HfNivel4"
      Tab(5).Control(5)=   "Text4"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Nivel 3"
      TabPicture(6)   =   "FrmCategoria.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Shape7"
      Tab(6).Control(1)=   "Shape6"
      Tab(6).Control(2)=   "Label5"
      Tab(6).Control(3)=   "Label6"
      Tab(6).Control(4)=   "HfNivel3"
      Tab(6).Control(5)=   "Text3"
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Nivel 2"
      TabPicture(7)   =   "FrmCategoria.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Command2"
      Tab(7).Control(1)=   "TxtIdNivel1"
      Tab(7).Control(2)=   "Txtcategoria2"
      Tab(7).Control(3)=   "cmdAgregar2"
      Tab(7).Control(4)=   "TxtInglesCategoria2"
      Tab(7).Control(5)=   "HfNivel2"
      Tab(7).Control(6)=   "Label4"
      Tab(7).Control(7)=   "Label3"
      Tab(7).Control(8)=   "Shape4"
      Tab(7).Control(9)=   "Shape5"
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "Nivel 1"
      TabPicture(8)   =   "FrmCategoria.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Command1"
      Tab(8).Control(1)=   "txtid_categoria"
      Tab(8).Control(2)=   "TxtCategoria2Ingles"
      Tab(8).Control(3)=   "cmdAgregar"
      Tab(8).Control(4)=   "TxtCategoria1"
      Tab(8).Control(5)=   "HfNivel1"
      Tab(8).Control(6)=   "Label2"
      Tab(8).Control(7)=   "Label1"
      Tab(8).Control(8)=   "Shape2"
      Tab(8).Control(9)=   "Shape3"
      Tab(8).ControlCount=   10
      TabCaption(9)   =   "Categoia"
      TabPicture(9)   =   "FrmCategoria.frx":00FC
      Tab(9).ControlEnabled=   -1  'True
      Tab(9).Control(0)=   "ShpDatos"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Shape1"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "LblEmpresa"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "LblFecha"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).Control(4)=   "HfCategoria"
      Tab(9).Control(4).Enabled=   0   'False
      Tab(9).Control(5)=   "Txtcategoria"
      Tab(9).Control(5).Enabled=   0   'False
      Tab(9).ControlCount=   6
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
         Height          =   285
         Left            =   -69000
         TabIndex        =   50
         Top             =   8520
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar"
         Height          =   285
         Left            =   -68880
         TabIndex        =   49
         Top             =   8640
         Width           =   1335
      End
      Begin VB.TextBox txtid_categoria 
         Height          =   375
         Left            =   -67320
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   8400
         Width           =   1695
      End
      Begin VB.TextBox TxtIdNivel1 
         Height          =   375
         Left            =   -67200
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   8280
         Width           =   1695
      End
      Begin VB.TextBox Txtcategoria2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72120
         TabIndex        =   43
         Top             =   8160
         Width           =   3015
      End
      Begin VB.CommandButton cmdAgregar2 
         Caption         =   "Agregar"
         Height          =   285
         Left            =   -69000
         TabIndex        =   42
         Top             =   8235
         Width           =   1335
      End
      Begin VB.TextBox TxtInglesCategoria2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72120
         TabIndex        =   41
         Top             =   8475
         Width           =   3015
      End
      Begin VB.TextBox TxtCategoria2Ingles 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   40
         Top             =   8520
         Width           =   3015
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   285
         Left            =   -68880
         TabIndex        =   39
         Top             =   8280
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   35
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   31
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   27
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   23
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   19
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   15
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   11
         Top             =   8400
         Width           =   3015
      End
      Begin VB.TextBox TxtCategoria1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -72000
         TabIndex        =   5
         Top             =   8200
         Width           =   3015
      End
      Begin VB.TextBox Txtcategoria 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         Top             =   8400
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCategoria 
         Height          =   7335
         Left            =   1680
         TabIndex        =   2
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel1 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   6
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel2 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   9
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel3 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   12
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel4 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   16
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel5 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   20
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   45.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel6 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   24
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel7 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   28
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel8 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   32
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfNivel9 
         Height          =   7335
         Left            =   -73320
         TabIndex        =   36
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12938
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73320
         TabIndex        =   44
         Top             =   8235
         Width           =   1035
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   38
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 9"
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
         Left            =   -73035
         TabIndex        =   37
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   34
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 8"
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
         Left            =   -73035
         TabIndex        =   33
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   30
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 7"
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
         Left            =   -73035
         TabIndex        =   29
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   26
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 6"
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
         Left            =   -73035
         TabIndex        =   25
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   22
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 5"
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
         Left            =   -73035
         TabIndex        =   21
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   18
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 4"
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
         Left            =   -73035
         TabIndex        =   17
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   14
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 3"
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
         Left            =   -73035
         TabIndex        =   13
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 2"
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
         Left            =   -73155
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   11415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -73200
         TabIndex        =   8
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 1"
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
         Left            =   -73035
         TabIndex        =   7
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   -73560
         Top             =   240
         Width           =   10335
      End
      Begin VB.Label LblFecha 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1800
         TabIndex        =   4
         Top             =   8280
         Width           =   1035
      End
      Begin VB.Label LblEmpresa 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORIAS"
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
         Left            =   1740
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   8775
         Left            =   1440
         Top             =   240
         Width           =   10335
      End
      Begin VB.Shape ShpDatos 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   1680
         Top             =   8160
         Width           =   9975
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   675
         Left            =   -73320
         Top             =   8160
         Width           =   5895
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   -73320
         Top             =   8090
         Width           =   5895
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":0118
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":056C
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":088C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":0CE0
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":1134
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":1454
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":1774
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":1A94
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCategoria.frx":1DB4
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3945
      Left            =   12120
      TabIndex        =   45
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6959
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3945
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   46
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
End
Attribute VB_Name = "FrmCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdAgregar_Click()

    If Val(Me.txtid_categoria.Text) = 0 Then
    strCadena = "INSERT INTO  categoria_1(id_categoria,descripcion,ingles) VALUES('" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ,'" & Trim(Me.TxtCategoria1.Text) & "','" & Trim(Me.TxtCategoria2Ingles.Text) & "')"
    CnBd.Execute (strCadena)
     
    Me.TxtCategoria1.Text = ""
    Me.TxtCategoria2Ingles.Text = ""
    
    Else
        strCadena = "update categoria_1 SET descripcion='" & Trim(Me.TxtCategoria1.Text) & "',ingles='" & Trim(Me.TxtCategoria2Ingles.Text) & "' WHERE id_categoria1='" & Val(Me.txtid_categoria.Text) & "'"
        CnBd.Execute (strCadena)
         
    End If
    
    strCadena = "SELECT id_categoria1 as id_codigo,descripcion,ingles  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "'  ORDER BY descripcion ASC"
    Call llenarGrid(Me.HfNivel1, strCadena)
End Sub

Private Sub cmdAgregar2_Click()
    If Val(Me.TxtIdNivel1.Text) = 0 Then
    strCadena = "INSERT INTO  categoria_2(id_categoria1,descripcion,ingles) VALUES('" & Val(Me.TXTIDentificador.Text) & "' ,'" & Trim(Me.Txtcategoria2.Text) & "','" & Trim(Me.TxtInglesCategoria2.Text) & "')"
    CnBd.Execute (strCadena)
     
    Me.Txtcategoria2.Text = ""
    Me.TxtInglesCategoria2.Text = ""
    
    Else
        strCadena = "update categoria_2 SET descripcion='" & Trim(Me.Txtcategoria2.Text) & "',ingles='" & Trim(Me.TxtInglesCategoria2.Text) & "' WHERE id_categoria2='" & Val(Me.TxtIdNivel1.Text) & "'"
        CnBd.Execute (strCadena)
         
    End If
    
    strCadena = "SELECT id_categoria2 as id_codigo,descripcion,ingles  FROM categoria_2 WHERE id_categoria1='" & Val(Me.TXTIDentificador.Text) & "'  ORDER BY descripcion ASC"
    Call llenarGrid(Me.HfNivel2, strCadena)
End Sub

Private Sub Command1_Click()
strCadena = "delete from  categoria_1  WHERE id_categoria1='" & Val(Me.txtid_categoria.Text) & "'"
CnBd.Execute (strCadena)
 
strCadena = "SELECT id_categoria1 as id_codigo,descripcion,ingles  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "'  ORDER BY descripcion ASC"
    Call llenarGrid(Me.HfNivel1, strCadena)
End Sub

Private Sub Command2_Click()
    strCadena = "delete from  categoria_2  WHERE id_categoria2='" & Val(Me.TxtIdNivel1.Text) & "'"
    CnBd.Execute (strCadena)
     

    strCadena = "SELECT id_categoria2 as id_codigo,descripcion,ingles  FROM categoria_2 WHERE id_categoria1='" & Val(Me.TxtIdNivel1.Text) & "'  ORDER BY descripcion ASC"
    Call llenarGrid(Me.HfNivel2, strCadena)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Call actualizar
End Sub
Public Sub actualizar()
strCadena = "SELECT id_categoria as id_codigo,descripcion,ingles  FROM categoria ORDER BY descripcion ASC"
Call llenarGrid(Me.HfCategoria, strCadena)

End Sub

Private Sub HfNivel1_Click()

Me.TXTIDentificador.Text = Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 0)
Me.txtid_categoria.Text = Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 0)
Me.TxtCategoria1.Text = Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 1)
Me.TxtCategoria2Ingles.Text = Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 2)
Me.CmdAgregar.Visible = True

End Sub

Private Sub HfNivel2_Click()
Me.TxtIdNivel1.Text = Me.HfNivel2.TextMatrix(Me.HfNivel2.Row, 0)
Me.Txtcategoria2.Text = Me.HfNivel2.TextMatrix(Me.HfNivel2.Row, 1)
Me.TxtInglesCategoria2.Text = Me.HfNivel2.TextMatrix(Me.HfNivel2.Row, 2)
Me.cmdAgregar2.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case Me.SSTab1.Tab
    Case 8
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion,ingles  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel1, strCadena)
    Case 7
         Me.TxtIdNivel1.Text = Val(Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 0))
        strCadena = "SELECT id_categoria2 as id_codigo,descripcion ,ingles FROM categoria_2 WHERE id_categoria1='" & Val(Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel2, strCadena)
    Case 6
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel3, strCadena)
    Case 5
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel4, strCadena)
    Case 4
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel5, strCadena)
    Case 3
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel6, strCadena)
    Case 2
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel7, strCadena)
    Case 1
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel8, strCadena)
    Case 0
        strCadena = "SELECT id_categoria1 as id_codigo,descripcion  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' ORDER BY descripcion ASC"
        Call llenarGrid(Me.HfNivel9, strCadena)
    
End Select
End Sub

Private Sub Text2_Change()

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_EXIT
         Unload Me
End Select
End Sub

Private Sub Toolbar9_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_EXIT
         Unload Me
         Exit Sub
End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
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
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 4000
           
        Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "INGLES"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = formato_item(rst("id_codigo"), 5) & vbTab & UCase(rst("descripcion")) & vbTab & UCase(rst("ingles"))
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

Private Sub TxtLinea_Change()

End Sub

Private Sub Txt_Change()


End Sub

Private Sub txtCategoria_Change()
strCadena = "SELECT id_categoria as id_codigo,descripcion,ingles  FROM categoria WHERE descripcion LIKE '%" & Trim(Me.Txtcategoria.Text) & "%' ORDER BY descripcion ASC"
Call llenarGrid(Me.HfCategoria, strCadena)
End Sub

Private Sub TxtCategoria1_Change()
strCadena = "SELECT id_categoria1 as id_codigo,descripcion,ingles  FROM categoria_1 WHERE id_categoria='" & Val(Me.HfCategoria.TextMatrix(Me.HfCategoria.Row, 0)) & "' AND descripcion LIKE '%" & Trim(Me.TxtCategoria1.Text) & "%' ORDER BY descripcion ASC"
Call llenarGrid(Me.HfNivel1, strCadena)
If Me.HfNivel1.Rows < 2 Then
    Me.CmdAgregar.Visible = True
Else
    Me.CmdAgregar.Visible = False
End If
    
    
End Sub

Private Sub Txtcategoria2_Change()
strCadena = "SELECT id_categoria1 as id_codigo,descripcion,ingles  FROM categoria_2 WHERE id_categoria1='" & Val(Me.HfNivel1.TextMatrix(Me.HfNivel1.Row, 0)) & "' AND descripcion LIKE '%" & Trim(Me.Txtcategoria2.Text) & "%' ORDER BY descripcion ASC"
Call llenarGrid(Me.HfNivel1, strCadena)
If Me.HfNivel1.Rows < 2 Then
    Me.cmdAgregar2.Visible = True
Else
    Me.cmdAgregar2.Visible = False
End If

End Sub

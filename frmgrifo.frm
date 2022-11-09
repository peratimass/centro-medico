VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmgrifo 
   BorderStyle     =   0  'None
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19995
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   19995
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   855
      Left            =   18000
      TabIndex        =   31
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVO"
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
      MICON           =   "frmgrifo.frx":0000
      PICN            =   "frmgrifo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdproducto 
      Height          =   750
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   5280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "PRODUCTO 1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":046E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdComprobante 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Tag             =   "1"
      Top             =   240
      Width           =   1850
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "FACTURA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   17.25
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
      MICON           =   "frmgrifo.frx":048A
      PICN            =   "frmgrifo.frx":04A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TIPO COMPROBANTE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6840
      TabIndex        =   0
      Top             =   3480
      Width           =   13575
      Begin VB.CommandButton Command6 
         Caption         =   "RUC/DNI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         TabIndex        =   18
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6720
         TabIndex        =   17
         Top             =   240
         Width           =   6615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "MASTERCARD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         TabIndex        =   16
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "VISA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         TabIndex        =   15
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EFECTIVO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         TabIndex        =   14
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   2640
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   1440
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   2640
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   1440
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   2640
         TabIndex        =   6
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   1440
         TabIndex        =   5
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdNumero 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   4200
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1815
         Left            =   6720
         TabIndex        =   13
         Top             =   3120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3201
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
   End
   Begin VitekeySoft.ChameleonBtn cmdComprobante 
      Height          =   615
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Tag             =   "3"
      Top             =   240
      Width           =   1850
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "BOLETA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":2A8B
      PICN            =   "frmgrifo.frx":2AA7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdComprobante 
      Height          =   615
      Index           =   2
      Left            =   4080
      TabIndex        =   21
      Tag             =   "54"
      Top             =   240
      Width           =   1850
      _ExtentX        =   3254
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "RECIBO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":508C
      PICN            =   "frmgrifo.frx":50A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdisla 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "ISLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   20.25
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
      MICON           =   "frmgrifo.frx":768D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdisla 
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "ISLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   20.25
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
      MICON           =   "frmgrifo.frx":76A9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdisla 
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "ISLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   20.25
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
      MICON           =   "frmgrifo.frx":76C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdisla 
      Height          =   735
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "ISLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   20.25
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
      MICON           =   "frmgrifo.frx":76E1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdproducto 
      Height          =   750
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   6120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "PRODUCTO 1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":76FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdproducto 
      Height          =   750
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   6960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "PRODUCTO 1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":7719
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdproducto 
      Height          =   750
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   7800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1323
      BTYPE           =   5
      TX              =   "PRODUCTO 1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   18
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
      MICON           =   "frmgrifo.frx":7735
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
      Height          =   2775
      Left            =   6840
      TabIndex        =   30
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   4895
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin VB.Image imgclose 
      Height          =   240
      Left            =   19560
      Picture         =   "frmgrifo.frx":7751
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3615
      Left            =   120
      Top             =   5040
      Width           =   5895
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3735
      Left            =   120
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmgrifo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComprobante_Click(Index As Integer)
Call load_producto(Me.cmdisla(Index).Tag)
For i = 0 To Me.cmdComprobante.Count - 1
    Me.cmdComprobante(i).BackColor = &HFFFFFF
Next i
Me.cmdComprobante(Index).BackColor = &H80FF&
End Sub

Private Sub cmdisla_Click(Index As Integer)
Call load_producto(Me.cmdisla(Index).Tag)
For i = 0 To Me.cmdisla.Count - 1
    Me.cmdisla(i).BackColor = &HFFFFFF
Next i
Me.cmdisla(Index).BackColor = &H80FF&

End Sub

Private Sub cmdproducto_Click(Index As Integer)
For i = 0 To Me.cmdproducto.Count - 1
    Me.cmdproducto(i).BackColor = &HFFFFFF
Next i
Me.cmdproducto(Index).BackColor = &H80FF&
End Sub

Private Sub Form_Load()
CenterForm Me
'Call load_skin("in_skin")
Me.Top = 50
Me.Caption = KEY_EMPRESA

Call load_isla
End Sub
Private Sub load_isla()
strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "' and id_sucursal='0' ORDER BY id_alm"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   For i = 0 To Me.cmdisla.Count - 1
       Me.cmdisla(i).Visible = False
       Me.cmdisla(i).BackColor = &HFFFFFF
   Next i
   
   For i = 0 To rst.RecordCount - 1
           Me.cmdisla(i).Visible = True
           Me.cmdisla(i).Caption = rst("descripcion")
           Me.cmdisla(i).Tag = rst("id_alm")
           rst.MoveNext
   Next i
End If
End Sub

Private Sub load_producto(ByVal in_isla As String)
strCadena = "SELECT id_producto,nombre_prod,precio_venta FROM view_producto WHERE ruc='" & KEY_RUC & "' and id_alm='" & in_isla & "' ORDER BY id_producto"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   For i = 0 To Me.cmdisla.Count - 1
       Me.cmdproducto(i).Visible = False
       Me.cmdproducto(i).BackColor = &HFFFFFF
   Next i
   
   For i = 0 To rst.RecordCount - 1
           Me.cmdproducto(i).Visible = True
           Me.cmdproducto(i).Caption = rst("nombre_prod")
           Me.cmdproducto(i).Tag = rst("id_producto")
           rst.MoveNext
   Next i
End If
End Sub

Private Sub imgclose_Click()
Unload Me
End Sub

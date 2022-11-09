VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmLibro 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtContenido 
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
      Left            =   4125
      TabIndex        =   51
      Top             =   8800
      Width           =   2775
   End
   Begin VB.TextBox txtCodigoClasificacion 
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
      Left            =   1530
      TabIndex        =   49
      Top             =   8800
      Width           =   1095
   End
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   14880
      ScaleHeight     =   4665
      ScaleMode       =   0  'User
      ScaleWidth      =   5145
      TabIndex        =   40
      Top             =   4080
      Width           =   5175
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   4125
      TabIndex        =   34
      Top             =   8420
      Width           =   2775
   End
   Begin VB.TextBox TxtCod 
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
      Left            =   1530
      TabIndex        =   33
      Top             =   8420
      Width           =   1095
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "<<< ANTERIOR"
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
      Left            =   14880
      TabIndex        =   32
      Top             =   8880
      Width           =   1575
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "SIGUIENTE >>>"
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
      Left            =   18480
      TabIndex        =   31
      Top             =   8880
      Width           =   1455
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   12720
      TabIndex        =   29
      Top             =   8420
      Width           =   855
   End
   Begin VB.CheckBox chkLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AREA       :"
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
      Height          =   240
      Left            =   8040
      TabIndex        =   28
      Top             =   8490
      Width           =   975
   End
   Begin VB.Frame frm_ubicacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MATRIZ UBICACION FISICA"
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
      Height          =   2295
      Left            =   15360
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox Txt_y 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_x 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TxtAndamio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtPiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TxtSector 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632319
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASILLERO :"
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
         Left            =   195
         TabIndex        =   27
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANDAMIO :"
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
         Left            =   225
         TabIndex        =   26
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PISO :"
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
         Left            =   645
         TabIndex        =   25
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECTOR :"
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
         Left            =   405
         TabIndex        =   24
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN :"
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
         Left            =   255
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame framemayor 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   15360
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtpreciomayor 
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
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtprecioventa 
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
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VitekeySoft.ChameleonBtn cmdactualizar 
         Height          =   405
         Left            =   1920
         TabIndex        =   12
         Top             =   1320
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "ACTUALIZAR P.MAYOR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         MICON           =   "frmLibro.frx":0000
         PICN            =   "frmLibro.frx":001C
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
         Height          =   405
         Left            =   1920
         TabIndex        =   13
         Top             =   1800
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "CERRAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         MICON           =   "frmLibro.frx":3227
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
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO MAYOR  :"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO NORMAL  :"
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
         Left            =   135
         TabIndex        =   14
         Top             =   480
         Width           =   1395
      End
   End
   Begin VB.CheckBox chkmarca 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AUTOR   :"
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
      Height          =   240
      Left            =   8040
      TabIndex        =   4
      Top             =   8860
      Width           =   975
   End
   Begin VB.TextBox txtMarca 
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
      Left            =   12720
      TabIndex        =   3
      Top             =   8820
      Width           =   855
   End
   Begin VB.Frame frmcompatible 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIMILITUD CON"
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
      Height          =   3975
      Left            =   6000
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   7695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCompatible 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5530
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
      Begin VitekeySoft.ChameleonBtn cmdCerrar 
         Height          =   260
         Left            =   7320
         TabIndex        =   2
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   260
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         MICON           =   "frmLibro.frx":3243
         PICN            =   "frmLibro.frx":325F
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   13800
      TabIndex        =   5
      Top             =   4660
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "SALIR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":6113
      PICN            =   "frmLibro.frx":612F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   13800
      TabIndex        =   6
      Top             =   2010
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":651F
      PICN            =   "frmLibro.frx":653B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   13800
      TabIndex        =   7
      Top             =   1120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":8985
      PICN            =   "frmLibro.frx":89A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   13800
      TabIndex        =   8
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":8CBB
      PICN            =   "frmLibro.frx":8CD7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdStockBajo 
      Height          =   405
      Left            =   13800
      TabIndex        =   30
      ToolTipText     =   "Stock Bajo"
      Top             =   7560
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "ST(-)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "frmLibro.frx":9129
      PICN            =   "frmLibro.frx":9145
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":C350
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":C7A4
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":CAC4
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":CF18
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":D36C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":D68C
            Key             =   "(Actualizar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":D7E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLibro.frx":DB00
            Key             =   "(merma)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   8055
      Left            =   120
      TabIndex        =   35
      Top             =   240
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14208
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAlmacen 
      Height          =   2655
      Left            =   14880
      TabIndex        =   36
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
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
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   330
      Left            =   9120
      TabIndex        =   37
      Top             =   8415
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VitekeySoft.ChameleonBtn cmdReport 
      Height          =   405
      Left            =   13800
      TabIndex        =   38
      ToolTipText     =   "Reporte"
      Top             =   8040
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "REPOR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "frmLibro.frx":10CDA
      PICN            =   "frmLibro.frx":10CF6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   405
      Left            =   13800
      TabIndex        =   39
      ToolTipText     =   "Precio x Mayor"
      Top             =   7080
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "frmLibro.frx":10D83
      PICN            =   "frmLibro.frx":10D9F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdlogistica 
      Height          =   525
      Left            =   13800
      TabIndex        =   41
      ToolTipText     =   "Ubicacion"
      Top             =   8520
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   926
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "frmLibro.frx":11339
      PICN            =   "frmLibro.frx":11355
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcMarca 
      Height          =   330
      Left            =   9120
      TabIndex        =   42
      Top             =   8820
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VitekeySoft.ChameleonBtn cmdcompatibility 
      Height          =   855
      Left            =   13800
      TabIndex        =   43
      Top             =   2890
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "SIMILITUDES"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":13A72
      PICN            =   "frmLibro.frx":13A8E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCaracteristicas 
      Height          =   855
      Left            =   13800
      TabIndex        =   44
      Top             =   3780
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CARACTERIS"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLibro.frx":177CF
      PICN            =   "frmLibro.frx":177EB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTENIDO:"
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
      Left            =   3075
      TabIndex        =   52
      Top             =   8880
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASIFICACION:"
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
      Left            =   240
      TabIndex        =   50
      Top             =   8880
      Width           =   1125
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TITULO:"
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
      Left            =   3405
      TabIndex        =   48
      Top             =   8550
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO :"
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
      Left            =   735
      TabIndex        =   47
      Top             =   8505
      Width           =   645
   End
   Begin VB.Label lblFotos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   16800
      TabIndex        =   45
      Top             =   8880
      Width           =   1155
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   120
      Top             =   8340
      Width           =   13575
   End
   Begin VB.Label LblProducto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   14880
      TabIndex        =   46
      Top             =   240
      Width           =   5175
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
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim strLinea As Boolean
Dim strMostrarTodos As Boolean





Private Sub ChameleonBtn1_Click()
Me.framemayor.Visible = False
End Sub

Private Sub CmdActualizar_Click()

strCadena = "UPDATE almacen_producto set precio_venta='" & Val(Me.txtprecioventa.Text) & "' WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

 
strCadena = "DELETE FROM almacen_producto_precio WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'  AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

 
strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "','" & KEY_ALM & "','" & Val(Me.txtpreciomayor.Text) & "','1','1','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

 


Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 7) = Format(Val(Me.txtpreciomayor.Text), "###0.00")
Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8) = Format(Val(Me.txtprecioventa.Text), "###0.00")
Me.framemayor.Visible = False
End Sub
Public Sub actualizar_update(ByVal in_libro As String)
 strCadena = "SELECT * FROM view_libro WHERE id_libro='" & in_libro & "' and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
 Call llenarGrid(Me.HfdGrilla, strCadena)
 

End Sub
Private Sub CmdAnterior_Click()
If rstI.EOF = True Or rstI.BOF = True Then
    If rstI.RecordCount < 1 Then
        Exit Sub
    End If
    rstI.MoveFirst
Else
    rstI.MovePrevious
    If rstI.EOF = True Or rstI.BOF = True Then
        rstI.MoveLast
    End If
End If
If IsNull(rstI("foto")) = False And Len(rstI("foto")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstI("foto")) = True Then
        'Me.Image1.Visible = True
        Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstI("foto")))
    Else
        'Me.Image1 = Nothing
    End If
End If
End Sub

Private Sub CmdLinea_Click()
End Sub

Private Sub cmdCaracteristicas_Click()
Me.frmcompatible.Caption = "CARACTERISTICAS"
Call load_caracteristicas(Me.HfCompatible, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
frmcompatible.Visible = True

End Sub

Private Sub cmdcerrar_Click()
Me.frmcompatible.Visible = False
End Sub

Private Sub cmdcompatibility_Click()

Me.frmcompatible.Caption = "COMPATIBILIDAD"
Call load_compatibility(Me.HfCompatible, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
frmcompatible.Visible = True

End Sub

Private Sub cmddelete_Click()
 If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        Call disabled_form(Me)
        frmsegurity.Show
        Exit Sub
        
          
          
            Call ActualizarProd
          '  Call ActualizarAlm
        End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdlogistica_Click()
'strCadena = "SELECT * FROM almacen_producto a WHERE id_alm='" & KEY_ALM & "' and  precio_mayor>0 and  ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
 '   rst.MoveFirst
  '  For i = 0 To rst.RecordCount - 1
   '     strCadena = "SELECT * FROM almacen_producto_precio WHERE id_alm='" & KEY_ALM & "' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
    '    Call ConfiguraRstT(strCadena)
     '   If rstT.RecordCount < 1 Then
      '      strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & rst("id_producto") & "','" & KEY_ALM & "','" & Val(rst("precio_mayor")) & "','1','1','" & KEY_RUC & "')"
       '     CnBd.Execute (strCadena)

      '  End If
      '  rst.MoveNext
      '  DoEvents
   ' Next i
'End If
   
   If Me.frm_ubicacion.Visible = True Then
      Me.frm_ubicacion.Visible = False
   Else
    Me.frm_ubicacion.Visible = True
   End If
End Sub

Private Sub cmdmodificar_Click()
Me.txtprecioventa.Text = Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8))
Me.txtpreciomayor.Text = Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 7))
Me.framemayor.Visible = True
End Sub

Private Sub cmdnuevo_Click()
      Procedencia = nuevo
      frmLibroDetalle.Show
      Exit Sub
End Sub

Private Sub cmdReport_Click()
If Me.chkLinea.Value = 1 Then
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,precio_compra,precio_venta FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea='" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod"
Else
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,precio_compra,precio_venta FROM view_producto WHERE stock<=stock_minimo and ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY id_linea,nombre_prod"
End If
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptStockMinimo", , App.Path + "\Reportes\")

End Sub

Private Sub cmdsiguiente_Click()
If rstI.EOF = True Or rstI.BOF = True Then
    If rstI.RecordCount < 1 Then
        Exit Sub
    End If
    rstI.MoveFirst
Else
    rstI.MoveNext
    If rstI.EOF = True Or rstI.BOF = True Then
        rstI.MoveFirst
    End If
End If
If IsNull(rstI("foto")) = False And Len(rstI("foto")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstI("foto")) = True Then
        'Me.Image1.Visible = True
        Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstI("foto")))
    Else
        'Me.Image1 = Nothing
    End If
End If
End Sub



Private Sub cmdStockBajo_Click()
If strLinea = False Then
 If KEY_SKFACTURA = "no" Then
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 0,50 "
    Call Me.llenarGrid(Me.HfdGrilla, strCadena)
    Me.cmdReport.Visible = True
    Exit Sub
  Else
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 0,50 "
    Call Me.llenarGrid(Me.HfdGrilla, strCadena)
    Me.cmdReport.Visible = True
    Exit Sub
  End If

End If

If strLinea = True Then
 If KEY_SKFACTURA = "no" Then
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea = '" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod LIMIT 0,50 "
    Call Me.llenarGrid(Me.HfdGrilla, strCadena)
    Exit Sub
  Else
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea = '" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod LIMIT 0,50 "
     Call llenarGrid(Me.HfdGrilla, strCadena)
    Exit Sub
  End If
End If
End Sub

Private Sub cmdupdate_Click()
      Procedencia = modificar
      frmLibroDetalle.Show
End Sub

Private Sub DtcLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If Me.chkLinea.Value = 1 Then
        strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'  and titulo LIKE '%" & Trim(Me.txtproducto.Text) & "%' and  id_area ='" & Trim(Me.DtcLinea.BoundText) & "' "
      Else
        strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and  id_area ='" & Trim(Me.DtcLinea.BoundText) & "' "
      End If
      
      Call llenarGrid(Me.HfdGrilla, strCadena)
      Me.DtcLinea.SetFocus
End If
End Sub

Private Sub DtcMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Me.chkmarca.Value = 1 Then
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'   and titulo LIKE '%" & Trim(Me.txtproducto.Text) & "%' and  id_autor = '" & Me.DtcMarca.BoundText & "' ORDER BY titulo "
Else
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND    id_autor = '" & Me.DtcMarca.BoundText & "' ORDER BY titulo "
End If
Call llenarGrid(Me.HfdGrilla, strCadena)
Me.DtcMarca.SetFocus

    End If
End Sub

Private Sub Form_Activate()
Me.txtproducto.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 27) Then
    Unload Me
End If
If Shift = 2 And KeyCode = Asc("R") Then
    If MsgBox("QUIERE RECOMEDAR ESTE PRODUCTO", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Dim codigo_r As String
        strCadena = "SELECT * FROM Producto_Recomendado ORDER BY id_producto DESC"
        Call ConfiguraRst(strCadena)
        codigo_r = GeneraCodigo(4)
        Set rst = Nothing
        strCadena = "INSERT INTO Producto_Recomendado VALUES ('" & Trim(codigo_r) & "','" & Trim(Me.txtproducto.Text) & "') "
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Call Resalta(Me.txtproducto)
        Exit Sub
       End If
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
'Me.Image1.BordeEstilo = Borde4

  strLinea = False
  strMostrarTodos = False
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  
   strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE id_proveedor='si' and  ruc='" & KEY_RUC & "'  ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
  
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
'Me.Label2.Caption = KEY_EMPRESA
Call ActualizarProd
'Call ActualizarAlm
If KEY_CARGO = "00004" Then
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = False
    Me.cmddelete.Enabled = False
    Me.cmdcompatibility.Enabled = False
    
    
Else
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = True
    Me.cmddelete.Enabled = True
  End If

    
End Sub


Private Sub CargarLogo(ByVal cproducto As String, ByVal id_producto As String)
 On Error GoTo salir
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & cproducto) = True Then
       Me.Image1.Visible = True
       Me.lblFotos.Caption = "1"
       Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(cproducto))
       strCadena = "SELECT * FROM producto_foto WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRstI(strCadena)
       If rstI.RecordCount > 0 Then
            Me.CmdAnterior.Enabled = True
            Me.CmdSiguiente.Enabled = True
            Me.lblFotos.Caption = str(rstI.RecordCount)
       Else
            Me.CmdAnterior.Enabled = False
            Me.CmdSiguiente.Enabled = False
       End If
    Else
        Me.Image1.Visible = False
        Me.Image1.Picture = Nothing
    End If
Exit Sub
salir:
    Me.Image1.Picture = Nothing
    Exit Sub
End Sub

Private Sub HfCompatible_DblClick()
If Me.HfCompatible.Rows > 0 Then
    If Val(Me.HfCompatible.TextMatrix(Me.HfCompatible.Row, 0)) > 0 Then
        strLinea = False
        Me.txtproducto.Text = Trim(Me.HfCompatible.TextMatrix(Me.HfCompatible.Row, 1))
        Call busqueda
        Me.frmcompatible.Visible = False
    End If
End If
End Sub

Private Sub HfdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
        Me.HfdGrilla.AllowBigSelection = True
'        Me.HfdGrilla.SetFocus
End If
If KeyCode = vbKeyUp Then
     Me.HfdGrilla.AllowBigSelection = True
    
End If
End Sub

Private Sub HfdGrilla_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And FrmVentas.Procedencia = Selecionar Then
     
    strCadena = "SELECT * FROM view_producto_selec WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmVentas.codigoP = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
        FrmVentas.TxtCodProducto.Text = Trim(rst("id_producto"))
        FrmVentas.TxtDescripcionProducto.Text = rst("nombre_prod")
        
        If FrmVentas.DtcMoneda.BoundText = "00002" Then
            FrmVentas.txtPrecio.Text = Format(rst("precio_venta") / (FrmVentas.TxtTipoCambio.Text), "###0.00")
            FrmVentas.txtpreciooriginal.Text = Format(rst("precio_venta") / (FrmVentas.TxtTipoCambio.Text), "###0.00")
        Else
            FrmVentas.txtPrecio.Text = rst("precio_venta")
            FrmVentas.txtpreciooriginal.Text = rst("precio_venta")
        End If
        
        
        
        FrmVentas.txtServicio.Text = rst("servicio")
        FrmVentas.txt_tipo.Text = rst("id_tipo")
        FrmVentas.txtpeso.Text = rst("peso")
        
        If KEY_SKFACTURA = "si" Then
            'FrmVentas.LblUnidad.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 3)
        Else
            'FrmVentas.LblUnidad.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 5)
        End If
        FrmVentas.TxtIgv.Text = rst("id_igv")
        
        If Val(FrmVentas.txtcantidad.Text) <> 0 Then
            FrmVentas.txtcantidad.Text = Val(FrmVentas.txtcantidad.Text)
        Else
            FrmVentas.txtcantidad.Text = 1
        End If
        FrmVentas.chkPrecios.Enabled = True
        FrmVentas.chkPrecios.Value = 1
        
        Call FrmVentas.mostrar_precios
        
        
        If rst("stock") <= 0 And rst("id_tipo") = "01" Then
        '    strCadena = "SELECT * FROM producto WHERE id_relacionado='" & rst("id_producto") & "' AND ruc='" & KEY_RUC & "'"
        '   Call ConfiguraRstT(strCadena)
        '  If rstT.RecordCount < 1 Then
             
        '       MsgBox "PRODUCTO NO CUENTA CON STOCK", vbInformation, KEY_EMPRESA
        '       Call Resalta(FrmVentas.TxtCodProducto)
             
        ' End If
        ' Set rstT = Nothing
        End If
        
        
        
   If FrmVentas.OptAuto.Value = True Then
            Call FrmVentas.Agregar_directo
     Else
        FrmVentas.txtPrecio.Locked = False
        Call FrmVentas.Resalta(FrmVentas.txtcantidad)
        'Call FrmVentas.Resalta(FrmVentas.txtprecio)
        'Call FrmVentas.mostrar_precios
        End If
        FrmVentas.Procedencia = Neutro
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If

If FrmDetalleProducto.Procedencia = Selecionar Then
   FrmDetalleProducto.Procedencia = Neutro
   FrmDetalleProducto.txtCodCompatible.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmDetalleProducto.TxtCompatible.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If


If FrmComprasGastos.Procedencia = Selecionar Then
   FrmComprasGastos.Procedencia = Neutro
   FrmComprasGastos.txtcodigoprod.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmComprasGastos.txtproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If







If KeyAscii = 13 And FrmProductosRelacionados.Procedencia = relacionar Then
    FrmProductosRelacionados.Procedencia = Neutro
    strCadena = "SELECT id_producto,nombre_prod FROM producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'  "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductosRelacionados.txtCodSubProducto.Text = rst(0)
        FrmProductosRelacionados.txtDescripcionSubproducto.Text = rst(1)
        Call Resalta(FrmProductosRelacionados.txtcantidad)
        Unload Me
        
        Set rst = Nothing
    End If
    Exit Sub
End If

If KeyAscii = 13 And FrmProductoSubproducto.Procedencia = relacionar Then
    
        FrmProductoSubproducto.txtCodSubProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        FrmProductoSubproducto.txtDescripcionSubproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
        Call Resalta(FrmProductoSubproducto.txtcantidad)
        Unload Me
        FrmProductoSubproducto.Procedencia = Neutro
     
    
    Exit Sub
End If


If frmCorProcesos.Procedencia = seleccionar_soldadura Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfSoldaduraInsumo)
        frmCorProcesos.Procedencia = Neutro
        Unload Me
        
        Exit Sub
End If


If frmCorProcesos.Procedencia = seleccionar_ensamblaje Then
        frmCorProcesos.Procedencia = Neutro
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfEnsambladoInsumo)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If

If FrmDetalleLinea.Procedencia = seleccionar_insumo Then
   FrmDetalleLinea.txtid_insumo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmDetalleLinea.lblinsumo.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call Resalta(FrmDetalleLinea.txtcantidad)
   FrmDetalleLinea.Procedencia = Neutro
   Unload Me
   Exit Sub
End If

If frmmantenimientos.Procedencia = seleccionar_insumo Then
   frmmantenimientos.frminsumo.Visible = True
   frmmantenimientos.txtid_insumo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmmantenimientos.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call Resalta(frmmantenimientos.txtcantidad)
   frmmantenimientos.Procedencia = Neutro
   Unload Me
   Exit Sub
End If


If frmCorProcesos.Procedencia = seleccionar_tapiz Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfTapizadoInsumo)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If

If frmCorProcesos.Procedencia = seleccionar_otro Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfInsumoTercero)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If




If KeyAscii = 13 And FrmKardexdeProductos.Procedencia = buscar Then
        FrmKardexdeProductos.DtcProducto.BoundText = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        Call FrmKardexdeProductos.presionar
         
        Unload Me
        FrmKardexdeProductos.Procedencia = Neutro
    Exit Sub
End If
If KeyAscii = 13 And FrmGeneradorBarras.Procedencia = Selecionar Then
        FrmGeneradorBarras.txtcodigo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
       ' Call FrmGeneradorBarras.presionar(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         
        Unload Me
        FrmGeneradorBarras.Procedencia = Neutro
    Exit Sub
End If

If KeyAscii = 13 And FrmVentasPersonalizada.Procedencia = Selecionar Then
   FrmVentasPersonalizada.TxtCodProducto(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmVentasPersonalizada.txtDescripcion(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   FrmVentasPersonalizada.txtPrecio(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8)
   FrmVentasPersonalizada.TxtUnidad(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 5)
   Call Resalta(FrmVentasPersonalizada.txtcantidad(FrmVentasPersonalizada.numeroItem))
   FrmVentasPersonalizada.Procedencia = Neutro
   Unload Me
   Exit Sub
End If





If frmVentasPagos.Procedencia = Selecionar Then
   frmVentasPagos.txtid_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmVentasPagos.txtObservacion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1) & "-" & "COLOR :" & "- " & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 4) & "-" & Trim(frmVentasPagos.txtObservacion.Text) & Space(2)
   frmVentasPagos.framevehiculo.Visible = True
   frmVentasPagos.txtmontovehiculo.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.TxtMontoReal.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.txtsaldo.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.Procedencia = Neutro
   Unload Me
   Exit Sub
End If

If KeyAscii = 13 And FrmBusquedaDocumentos.Procedencia = Selecionar Then
    FrmBusquedaDocumentos.TxtCodigoInterno.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmBusquedaDocumentos.txtCodigoProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmBusquedaDocumentos.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 2)
    FrmBusquedaDocumentos.txtDescripcion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
    FrmBusquedaDocumentos.cmdBuscar_producto.Enabled = True
    FrmBusquedaDocumentos.cmdBuscar_producto.SetFocus
    FrmBusquedaDocumentos.Procedencia = Neutro
    Exit Sub
End If


If KeyAscii = 13 And FrmDetalleLinea.Procedencia = Selecionar Then
    FrmDetalleLinea.txtid_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmDetalleLinea.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
    FrmDetalleLinea.lblcosto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 7)
    
    FrmDetalleLinea.Procedencia = Neutro
    Unload Me
    Exit Sub
End If

'****************MERMAS******************
If KeyAscii = 13 And FrmProductoMermas.Procedencia = mermas Then
     strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoMermas.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoMermas.cod_producto = rst("id_producto")
        FrmProductoMermas.txtid_producto.Text = rst("id_producto")
        FrmProductoMermas.txtproducto.Text = rst("nombre_prod")
        FrmProductoMermas.lblStock.Caption = rst("stock")
        FrmProductoMermas.txtcosto.Text = rst("precio_compra")
        'FrmProductoMermas.DtcMerma.SetFocus
        FrmProductoMermas.Procedencia = Neutro
        Unload Me
       
        Set rst = Nothing
    End If
    Exit Sub
End If


If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones And FrmProductoTransformaciones.prodA = True Then
    strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.txtcodigo.Text = rst("id_producto")
        FrmProductoTransformaciones.txtproducto.Text = rst("nombre_prod")
        FrmProductoTransformaciones.lblStock.Caption = rst("stock")
        FrmProductoTransformaciones.txtcosto.Text = rst("precio_compra")
        FrmProductoTransformaciones.LblUnidad.Caption = rst("abreviatura")
        FrmProductoTransformaciones.TxtPVenta.Text = rst("precio_venta")
        Unload Me
        Set rst = Nothing
    End If
    FrmProductoTransformaciones.Procedencia = Neutro
    FrmProductoTransformaciones.prodA = False
    Exit Sub
End If
If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones And FrmProductoTransformaciones.prodB = True Then
    strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.TxtCodigobarraB.Text = rst("id_producto")
        FrmProductoTransformaciones.txtDescripcionB.Text = rst("nombre_prod")
        FrmProductoTransformaciones.LblStockB.Caption = rst("stock")
        FrmProductoTransformaciones.txtCostoB.Text = rst("precio_compra")
        FrmProductoTransformaciones.lblUndB.Caption = rst("abreviatura")
        FrmProductoTransformaciones.TxtVentaB.Text = rst("precio_venta")
        Unload Me
        Set rst = Nothing
    End If
     FrmProductoTransformaciones.Procedencia = Neutro
    FrmProductoTransformaciones.prodB = False
    Exit Sub
End If






If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones Then
    Me.HfdGrilla.col = 0
    
   strCadena = "SELECT Producto_barras.cProducto, Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
    "Almacen_Productos.Stock,Producto.PrecioCompra,Producto.PrecioVenta FROM Producto_barras INNER JOIN Producto ON Producto_barras.cProducto = Producto.cProducto INNER JOIN " & _
    "Almacen_Productos ON Producto_barras.cProducto = Almacen_Productos.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
    "WHERE Almacen_Productos.cProducto='" & Trim(Me.HfdGrilla.Text) & "' AND Almacen_Productos.Alm_cod='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.cod_producto = rst(0)
        FrmProductoTransformaciones.txtcodigo.Text = rst(1)
        FrmProductoTransformaciones.txtproducto.Text = rst(2)
        FrmProductoTransformaciones.lblStock.Caption = rst(4)
        FrmProductoTransformaciones.txtcosto.Text = rst(5)
        FrmProductoTransformaciones.LblUnidad.Caption = rst("sAbreviatura")
        FrmProductoTransformaciones.TxtPVenta.Text = rst("PrecioVenta")
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If

If KeyAscii = 13 And FrmInventario.Procedencia = Selecionar Then
       strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "'  and id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_alm='" & Trim(FrmInventario.DtcAlmacen.BoundText) & "'"
       Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         FrmInventario.TxtCodProducto.Text = rst("id_producto")
         FrmInventario.txtid_producto.Text = rst("id_producto")
         FrmInventario.TxtDescripcionProducto.Text = rst("nombre_prod")
         FrmInventario.TxtStck_actual.Text = rst("stock")
         FrmInventario.txtStock_factura.Text = rst("stock_factura")
         FrmInventario.TxtUnidad.Text = rst("unidad")
         FrmInventario.TxtVenta.Text = rst("precio_venta")
         FrmInventario.txtcosto.Text = rst("precio_compra")
         FrmInventario.DtcClasificacion.BoundText = rst("id_linea")
         FrmInventario.DtcModelo.BoundText = rst("id_sublinea")
         FrmInventario.cmdStock.Enabled = True
         If rst("produccion") = "si" Then
            FrmInventario.cmdSeriales.Visible = True
         Else
            FrmInventario.cmdSeriales.Visible = False
         End If
         Call FrmInventario.Resalta(FrmInventario.TxtStock_nuevo)
        
         Set rst = Nothing
    End If
        FrmInventario.Procedencia = Neutro
         Unload Me
    Exit Sub
End If


If KeyAscii = 13 And FrmCompras.Procedencia = Selecionar Then
   FrmCompras.Procedencia = Neutro
   
    strCadena = "SELECT A.id_producto,P.nombre_prod,U.descripcion as abreviatura,A.stock,A.precio_compra,A.precio_venta,P.id_linea,P.numero_procedimientos,P.agranel FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND A.id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         FrmCompras.codigoP = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         FrmCompras.TxtCodProducto.Text = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         FrmCompras.TxtDescripcionProducto.Text = UCase(rst("nombre_prod")) & Space(5) & "[" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 6)) & "]"
         FrmCompras.txtcosto.Text = rst("precio_compra")
         FrmCompras.TxtUnidades.Text = rst("numero_procedimientos")
         Call FrmCompras.get_unidad(rst("id_producto"), rst("agranel"))
         
         FrmCompras.txtcantidad.Text = 0
         FrmCompras.TxtCostoAnt.Text = rst("precio_compra")
         FrmCompras.txtPrecioVentaAnt.Text = rst("precio_venta")
         FrmCompras.TxtventaHoy.Text = rst("precio_venta")
         
         If get_produccion(rst("id_linea")) = True Then
            FrmCompras.txtvalidacion_chasis.Text = "si"
         Else
            FrmCompras.txtvalidacion_chasis.Text = "no"
         End If
         
         
         Dim utilidad As Single
         If rst("precio_compra") > 0 Then
            utilidad = (rst("precio_venta") - rst("precio_compra")) * 100 / rst("precio_compra")
            FrmCompras.TxtUtilidadAnt.Text = Format(utilidad, "#,##0.00")
         End If
        Call Resalta(FrmCompras.txtcantidad)
        
        Unload Me
        
    End If
    Exit Sub
End If


If KeyAscii = 13 And FrmModificarCompras.Procedencia = Selecionar Then
    Me.HfdGrilla.col = 0
    strCadena = "SELECT Producto_barras.cProducto,Producto_barras.cod_barra, Producto.DescripcionProducto, Producto.PrecioCompra, Unidad.sAbreviatura,Producto.PrecioVenta " & _
    "FROM Producto INNER JOIN Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN " & _
    "Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE Producto.cProducto='" & Trim(Me.HfdGrilla.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmModificarCompras.codigoP = Trim(rst(0))
        FrmModificarCompras.TxtCodProducto.Text = rst(1)
        FrmModificarCompras.TxtDescripcionProducto.Text = rst(2)
        FrmModificarCompras.TxtUnidad.Text = rst(4)
        Unload Me
        Set rst = Nothing
    End If
    FrmModificarCompras.Procedencia = Neutro
    Exit Sub
End If
If KeyAscii = 13 And FrmTransferencias.Procedencia = Selecionar Then
        FrmTransferencias.Procedencia = Neutro
        strCadena = "SELECT * FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            FrmTransferencias.TxtCodProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmTransferencias.cprod = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmTransferencias.TxtDescripcionProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
            FrmTransferencias.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 5)
            FrmTransferencias.txtpeso.Text = rst("peso")
            FrmTransferencias.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 3)
            FrmTransferencias.txtcantidad.Enabled = True
            FrmTransferencias.txtcantidad.Text = ""
            Call Resalta(FrmTransferencias.txtcantidad)
            Unload Me
            Set rst = Nothing
    End If
    
    Exit Sub
End If

If KeyAscii = 13 And FrmParteMaterial.Procedencia = Selecionar Then
        strCadena = "SELECT * FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            FrmParteMaterial.TxtCodProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmParteMaterial.cprod = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmParteMaterial.TxtDescripcionProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
            FrmParteMaterial.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 2)
            FrmParteMaterial.txtpeso.Text = rst("peso")
            FrmParteMaterial.txtcantidad.Text = ""
            Call Resalta(FrmParteMaterial.txtcantidad)
            Unload Me
            Set rst = Nothing
    End If
    FrmParteMaterial.Procedencia = Neutro
    Exit Sub
End If


End Sub
Private Function get_produccion(ByVal in_linea As String) As Boolean
strCadena = "SELECT produccion FROM linea WHERE id_linea='" & in_linea & "' and id_usu='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   If rstK("produccion") = "si" Then
      get_produccion = True
   Else
      get_produccion = False
   End If
   
End If


End Function
Private Sub HfdGrilla_SelChange()
  
  strCadena = "SELECT * FROM view_libro_almacen WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
  Call ConfiguraRstP(strCadena)
  If rstP.RecordCount > 0 Then
     rstP.MoveFirst
     Me.lblproducto.Caption = rstP.Fields("nombre_prod")
     img = rstP("imagen")
     
     Call ActualizarAlm(Me.HfgAlmacen, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
    If KEY_FOTO = "si" And Len(img) > 0 Then
        Me.Image1.Visible = True
        Call CargarLogo(img, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
    
        Else
        Me.Image1.Visible = False
    End If
If KEY_CARGO = "00001" Or KEY_CARGO = "00009" Or KEY_CARGO = "00004" Then
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = True
    Me.cmddelete.Enabled = True
    Me.cmdcompatibility.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
Else
    Me.cmdNuevo.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmddelete.Enabled = False
    Me.cmdcompatibility.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
End If
Else
    Me.HfgAlmacen.Rows = 0
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = False
    Me.cmddelete.Enabled = False
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
End If
End Sub


Private Sub ActualizaStock()
Dim Producto As String
Dim Total As Double
Dim Registros As Integer
Dim codalmacen As String
Dim TotalProd As Double
Dim i As Integer
strCadena = "Select * from  producto "
Call ConfiguraRst(strCadena)
Registros = rst.RecordCount
Set rst = Nothing
'producto = "000"
codalmacen = "0001"
For i = 1 To Registros
    Producto = FormatosCeros(str(i), 4)
    
    strCadena = "SELECT Stk_Cant FROM Kardex WHERE cProducto='" & Trim(Producto) & "' AND Alm_Cod='" & codalmacen & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "SELECT SUM(Stk_Cant) FROM Kardex WHERE cProducto='" & Trim(Producto) & "' AND Alm_Cod='" & codalmacen & "'"
        Call ConfiguraRst(strCadena)
        TotalProd = rst(0)
        Set rst = Nothing
    Else
        GoTo salir
    End If
    
       
    strCadena = "UPDATE Almacen_Productos SET Stock='" & TotalProd & "' WHERE cProducto='" & Trim(Producto) & "' AND Alm_Cod='" & Trim(codalmacen) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    
    strCadena = "SELECT SUM(stock) FROM Almacen_Productos WHERE cProducto='" & Trim(Producto) & "' "
    Call ConfiguraRst(strCadena)
    Total = rst(0)
    Set rst = Nothing
    
    strCadena = "UPDATE Producto SET StockActual='" & Total & "' WHERE cProducto='" & Trim(Producto) & "' "
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
salir:

Next i

End Sub

Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
Dim X As Integer
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Rows = 1
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 600
  Grilla.ColWidth(1) = 5500
  Grilla.ColWidth(2) = 700
  Grilla.ColWidth(3) = 1100
  Grilla.ColAlignment(3) = 7
  Grilla.ColWidth(4) = 1100
  Grilla.ColAlignment(4) = 7
  Grilla.ColWidth(5) = 1100
  Grilla.ColAlignment(5) = 7
  Formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
Dim in_precio As Single
On Error GoTo salir

Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmddelete.Enabled = False
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 800
                Grilla.ColWidth(1) = 1200
                Grilla.ColWidth(2) = 5000
                Grilla.ColWidth(3) = 3000
                Grilla.ColWidth(4) = 1800
                Grilla.ColWidth(5) = 1500
                Grilla.ColWidth(6) = 1000
                
            Next
                cabecera = "CODIGO" & vbTab & "COD CLASIFICACION" & vbTab & "TITULO" & vbTab & "AUTOR" & vbTab & "AREA" & vbTab & "EDITORIAL" & vbTab & "TIPO LIBRO"
                Grilla.AddItem cabecera
                For k = 0 To 6
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
      
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
           
                Fila = rst("id_libro") & vbTab & rst("codigo_libro") & vbTab & rst("titulo") & vbTab & rst("autor") & vbTab & rst("area") & vbTab & rst("editorial") & vbTab & rst("tipo_libro")
                Grilla.AddItem Fila
            
            rst.MoveNext
    Next i
  Grilla.ColAlignment(1) = 1
  Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub llenarGridColor(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 900
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "STOCK" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("stock"), "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(rst("precio_venta"), "#,##0.00")
            If (Fila = "") Then
                X = 1
            End If
          Grilla.AddItem Fila
                        
                        If (Trim(rst("Stock")) < 2) Then
                            For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                      End If
            Fila = ""
            rst.MoveNext
             
        Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub LlenarGrid_Factura(ByVal Grilla As MSHFlexGrid, ByVal sql As String)

On Error GoTo salir
Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 1900
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 700
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 1000
           Grilla.ColWidth(8) = 1000
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "STOCK LOCAL" & vbTab & "STOCK FACTURA" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
        Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & UCase(rst("linea")) & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & rst("stock") & vbTab & rst("stock_factura") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(rst("precio_venta"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub



Public Sub LlenarGridcolorFactura(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
On Error GoTo salir


Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 4800
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1000
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "STOCK" & vbTab & "S-FACTURA" & vbTab & "P.VENTA" & vbTab & "P.COSTO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & rst("stock") & vbTab & rst("stock_factura") & vbTab & rst("precio_compra") & vbTab & rst("precio_venta")
           
          Grilla.AddItem Fila
                        
                        If (Trim(rst("stock")) < 2) Then
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                      End If
            Fila = ""
            rst.MoveNext
             
        Next i
        
   Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub actualizar()
    Call ActualizarProd
    'Call ActualizarAlm
End Sub
Public Sub ActualizarProd()
If KEY_ALM = "" Then
   KEY_ALM = "00001"
End If
 
 strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "'  LIMIT 30 "
 Call llenarGrid(Me.HfdGrilla, strCadena)
 Exit Sub
 
      
      
 

End Sub


Public Sub ActualizarAlm1()
If Val(Me.HfdGrilla.Rows) > 0 Then

  strCadena = "SELECT A.id_alm as CODIGO,A.descripcion AS ALMACEN,AP.stock AS Stock_Fisico,AP.stock_contable as Stock_Cont FROM almacen_producto AP,almacen A WHERE AP.id_alm=A.id_alm AND AP.id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND AP.ruc='" & KEY_RUC & "' AND A.ruc='" & KEY_RUC & "' ORDER BY A.stock DESC"
  Call ConfiguraRst(strCadena)
  HfgAlmacen.Clear
  'HfgAlmacen.Rows = 0
  Set Me.HfgAlmacen.Recordset = rst
  Me.HfgAlmacen.Rows = rst.RecordCount + 1
  Me.HfgAlmacen.ColWidth(0) = 0
  Me.HfgAlmacen.ColWidth(1) = 2500
  Me.HfgAlmacen.ColWidth(2) = 1100
  Me.HfgAlmacen.ColWidth(3) = 1100
  
End If
End Sub
Public Sub ActualizarAlm(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
'strCadena = "SELECT A.id_alm ,A.descripcion ,AP.stock,AP.stock_contable FROM almacen_producto AP,almacen A WHERE AP.id_alm=A.id_alm AND AP.id_producto='" & id_producto & "' AND AP.ruc='" & KEY_RUC & "' AND A.ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
If rstP.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstP.Fields.Count)
       For Each Campo In rstP.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2100
           Grilla.ColWidth(2) = 900
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 900
           
         Next
        cabecera = "CODIGO" & vbTab & "BIBLIOTECA" & vbTab & "STOCK " & vbTab & "PRESTADO " & vbTab & "TOTAL "
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstP.MoveFirst
        For i = 0 To rstP.RecordCount - 1
            Fila = rstP("id_alm") & vbTab & rstP("descripcion") & vbTab & rstP("stock") & vbTab & rstP("stock_contable") & vbTab & rstP("stock") + rstP("stock_contable")
            Grilla.AddItem Fila
            Fila = ""
            If rstP("id_alm") = KEY_ALM Then
               For k = 1 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
                 Me.TxtSector.Text = rstP("sector")
                 Me.TxtPiso.Text = rstP("piso")
                 Me.TxtAndamio.Text = rstP("andamio")
                 Me.txt_x.Text = rstP("casillero_x")
                 Me.Txt_y.Text = rstP("casillero_y")
  
            End If
            rstP.MoveNext
    Next i
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstP = Nothing

End Sub
Public Sub load_compatibility(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_compatibilidad WHERE id_padre='" & Trim(id_producto) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 1500
           
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "STOCK "
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("stock")
            Grilla.AddItem Fila
             
            rst.MoveNext
    Next i
        
     
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstP = Nothing

End Sub

Public Sub load_caracteristicas(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
strCadena = "SELECT * FROM producto_caracteristicas WHERE id_producto='" & Trim(id_producto) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 5500
           
       Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("caracteristica")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
     
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"


End Sub

Sub llenarGrid_alm(ByVal Grilla As MSHFlexGrid, ByVal sql As String)

  Call ConfiguraRst(sql)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 3900
  Grilla.Enabled = False
Grilla.Refresh

End Sub

Private Sub txtBuscar_Change()
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' AND descripcion like '%" & Trim(Me.txtBuscar.Text) & "%' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  area LIKE '%" & Trim(Me.txtBuscar.Text) & "%' "
    Call llenarGrid(Me.HfdGrilla, strCadena)
    Me.DtcLinea.SetFocus
End If
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

Dim Registros As Integer
Dim Criterio As String
  If Me.TxtCod.Text = "" Then
    Call ActualizarProd
    Exit Sub
  Else
  If Len(Me.TxtCod.Text) > 0 Then
  If KEY_BARRAS = "no" Then
    Criterio = " id_libro LIKE '%" & Trim(Me.TxtCod.Text) & "%'"
  Else
  Criterio = "B.cod_barra= '" & Trim(Me.TxtCod.Text) & "'"
  End If
  
   If KEY_SKFACTURA = "no" Then
      If KEY_BARRAS = "si" Then
        strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
        " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 0,100"
 
     Else
        strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY titulo LIMIT 0,50 "
        Call llenarGrid(Me.HfdGrilla, strCadena)
    End If
   
    Else
      If KEY_BARRAS = "si" Then
    strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,A.stock_factura,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
 " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 0,100"
 
    Else
        strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY titulo LIMIT 0,50 "
    End If
    Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
  End If
End If
  
    Me.TxtCod.SetFocus
  End If

End If
End Sub


Private Sub txtCodigoClasificacion_KeyPress(KeyAscii As Integer)
Dim Registros As Integer
Dim Criterio As String
If KeyAscii = 13 Then
  If Me.txtCodigoClasificacion.Text = "" Then
    Call ActualizarProd
    Exit Sub
  Else
  If Len(Me.txtCodigoClasificacion.Text) > 0 Then
  
    Criterio = " codigo_libro LIKE '%" & Trim(Me.txtCodigoClasificacion.Text) & "%'"
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY titulo LIMIT 0,50 "
    Call llenarGrid(Me.HfdGrilla, strCadena)
    Me.txtCodigoClasificacion.SetFocus
  End If
  End If
End If
End Sub

Private Sub txtContenido_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Call busqueda_contenido
End If
End Sub

Private Sub TxtMarca_Change()
  
  strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE id_proveedor='si' and  ruc='" & KEY_RUC & "' AND nombre_completo like '%" & Trim(Me.TxtMarca.Text) & "%' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)

End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  autor LIKE '%" & Trim(Me.TxtMarca.Text) & "%'"
    Call llenarGrid(Me.HfdGrilla, strCadena)
End If
End Sub

Private Sub TxtProducto_Change()
'Me.chkStockBajo.Value = 0

End Sub
Public Sub busqueda()
Dim parametros() As String
Dim Criterio As String

               
                
                    parametros = Split(Replace(Trim(Me.txtproducto.Text), "'", ""), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                
    
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  titulo LIKE '%" & Criterio & "%'  LIMIT 100 "
    Call llenarGrid(Me.HfdGrilla, strCadena)
    Me.txtproducto.SetFocus
End Sub
Public Sub busqueda_contenido()
Dim parametros() As String
Dim Criterio As String
                    parametros = Split(Replace(Trim(Me.TxtContenido.Text), "'", ""), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                
    strCadena = "SELECT * FROM view_libro WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  contenido LIKE '%" & Criterio & "%'  LIMIT 100"
    Call llenarGrid(Me.HfdGrilla, strCadena)
    Me.TxtContenido.SetFocus
    
End Sub

Private Sub TxtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        
        Me.HfdGrilla.SetFocus
    End If
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    strLinea = False
    Call busqueda
End If
End Sub


Private Sub Resalta(ByVal Texto As TextBox)
On Error GoTo Saltar
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
Saltar:
Exit Sub
End Sub



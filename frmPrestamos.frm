VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPrestamos 
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
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
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
      Height          =   7815
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   13815
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE CUENTA BANCARIA"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3495
         Left            =   480
         TabIndex        =   17
         Top             =   4200
         Width           =   12015
         Begin VB.TextBox TxtOperacion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2040
            TabIndex        =   23
            Top             =   1080
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DtcCuentaBancaria 
            Height          =   360
            Left            =   2040
            TabIndex        =   21
            Top             =   480
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcAutorizado 
            Height          =   360
            Left            =   2040
            TabIndex        =   25
            Top             =   2280
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdRegistrar 
            Height          =   780
            Left            =   8235
            TabIndex        =   26
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1376
            BTYPE           =   5
            TX              =   "REGISTRAR"
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmPrestamos.frx":0000
            PICN            =   "frmPrestamos.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdTransferir 
            Height          =   780
            Left            =   9450
            TabIndex        =   27
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1376
            BTYPE           =   5
            TX              =   "TRANSFERIR"
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmPrestamos.frx":3664
            PICN            =   "frmPrestamos.frx":3680
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpFechaDesembolso 
            Height          =   375
            Left            =   2040
            TabIndex        =   29
            Top             =   2880
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   52101121
            CurrentDate     =   44584
         End
         Begin MSDataListLib.DataCombo DtcFormaPago 
            Height          =   360
            Left            =   2040
            TabIndex        =   31
            Top             =   1680
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdImprimir 
            Height          =   780
            Left            =   10680
            TabIndex        =   34
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1376
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmPrestamos.frx":5F6A
            PICN            =   "frmPrestamos.frx":5F86
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AUTORIZADO POR :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   240
            TabIndex        =   30
            Top             =   2400
            Width           =   1395
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA DESEMBOLSO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   240
            TabIndex        =   28
            Top             =   3000
            Width           =   1665
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FORMA PAGO  :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   24
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° OPERACION :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   22
            Top             =   1200
            Width           =   1170
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CUENTA ORIGEN  :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE PRESTAMO"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3735
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   12015
         Begin VB.TextBox txtMonto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2040
            TabIndex        =   16
            Top             =   2520
            Width           =   2295
         End
         Begin VB.TextBox TxtDetalle 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1560
            Width           =   5535
         End
         Begin MSComCtl2.DTPicker DtpFecha 
            Height          =   375
            Left            =   2040
            TabIndex        =   13
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   52101121
            CurrentDate     =   44584
         End
         Begin MSDataListLib.DataCombo DtcPersonal 
            Height          =   360
            Left            =   2040
            TabIndex        =   14
            Top             =   960
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcPeriodo 
            Height          =   360
            Left            =   2040
            TabIndex        =   19
            Top             =   3120
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcMoneda 
            Height          =   360
            Left            =   4680
            TabIndex        =   32
            Top             =   2520
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERIODO PAGO  :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   360
            TabIndex        =   18
            Top             =   3120
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO PRESTAMO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   135
            TabIndex        =   12
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DETALLE PRESTAMO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   15
            TabIndex        =   11
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERSONAL   :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   630
            TabIndex        =   10
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA REGISTRO  :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   165
            TabIndex        =   9
            Top             =   480
            Width           =   1410
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdSalir 
         Height          =   250
         Left            =   13320
         TabIndex        =   35
         Top             =   240
         Width           =   250
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   ""
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
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrestamos.frx":8557
         PICN            =   "frmPrestamos.frx":8573
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18720
      TabIndex        =   0
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVA "
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrestamos.frx":B427
      PICN            =   "frmPrestamos.frx":B443
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   8175
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   14420
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12582912
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   975
      Left            =   18720
      TabIndex        =   2
      Top             =   2805
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "ANULAR"
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrestamos.frx":B895
      PICN            =   "frmPrestamos.frx":B8B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrar 
      Height          =   975
      Left            =   18720
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "CERRAR"
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrestamos.frx":DCFB
      PICN            =   "frmPrestamos.frx":DD17
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   975
      Left            =   18720
      TabIndex        =   33
      Top             =   1750
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "VISUALIZAR"
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrestamos.frx":10D3E
      PICN            =   "frmPrestamos.frx":10D5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRESTAMOS Y ADELANTOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   15600
      TabIndex        =   6
      Top             =   240
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   18255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdImprimir_Click()
Dim cam3(0 To 5, 1 To 5)  As String


    
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    cam3(0, 2) = Format(Me.DtpFecha.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFecha.Value, "dd-mm-YYYY")
    cam3(2, 2) = ""
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = ""
    param = cam3()
    strCadena = "call ADM_prestamo('5','" & Val(Me.frmdetalle.Tag) & "','','','','','','','','','','','','','','','','" & KEY_RUC & "')"
    
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptPrestamoPersonal", param, App.Path + "\Reportes\")
    Exit Sub
    
    
End Sub

Private Sub cmdnuevo_Click()
Call nuevo
End Sub

Private Sub cmdRegistrar_Click()

If Trim(Me.TxtOperacion.Text) <> "" And Trim(Me.TxtDetalle.Text) <> "" Then
   
    strCadena = "call ADM_prestamo('1','','" & KEY_ALM & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "' " & _
    ",'" & Me.DtcPersonal.BoundText & "','" & UCase(Me.TxtDetalle.Text) & "','" & Val(Me.txtMonto.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcPeriodo.BoundText & "', " & _
    "'" & Me.DtcCuentaBancaria.BoundText & "','" & Trim(Me.TxtOperacion.Text) & "','" & Me.DtcAutorizado.BoundText & "','si','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & Me.DtcFormaPago.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmdetalle.Visible = False
    Me.cmdRegistrar.Enabled = False
    Me.cmdTransferir.Enabled = True
    Call llenarGrid(Me.HfgDetalle)
    
Else
    MsgBox "INGRESE UNA OBSERVACION", vbInformation, KEY_VENDEDOR
End If

End Sub

Private Sub cmdSalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdTransferir_Click()

        
strCadena = "call ADM_prestamo('2','" & Me.frmdetalle.Tag & "','" & KEY_ALM & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaDesembolso.Value, "YYYY-mm-dd") & "','" & Trim(Me.DtcPersonal.BoundText) & "','" & Trim(Me.TxtDetalle.Text) & "','" & Val(Me.txtMonto.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & Me.DtcCuentaBancaria.BoundText & "','" & Trim(Me.TxtOperacion.Text) & "','" & Me.DtcAutorizado.BoundText & "','','" & KEY_CAMBIO_COMPRA & "','" & KEY_USUARIO & "','" & Me.DtcFormaPago.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.cmdTransferir.Enabled = False
End If


End Sub

Private Sub cmdVisualizar_Click()
strCadena = "call ADM_prestamo('4','" & Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0) & "','','','','','','','','','','','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.frmdetalle.Tag = rst("id")
    Me.DtpFecha.Value = rst("fecha_registro")
    Me.DtcPersonal.BoundText = rst("dni")
    Me.TxtDetalle.Text = rst("detalle")
    Me.txtMonto.Text = Format(rst("monto"), "#,##0.00")
    Me.DtcPeriodo.BoundText = rst("perido")
    Me.DtcCuentaBancaria.BoundText = rst("id_cuenta")
    Me.TxtOperacion.Text = rst("operacion")
    Me.DtcFormaPago.BoundText = rst("id_forma_pago")
    Me.DtcAutorizado.BoundText = rst("dni_autorizado")
    Me.DtpFechaDesembolso.Value = rst("fecha_transferencia")
    If rst("pendiente") = "si" Then
        Me.cmdTransferir.Enabled = True
    Else
        Me.cmdTransferir.Enabled = False
    End If
    
    Me.cmdRegistrar.Enabled = False
    
    Me.frmdetalle.Visible = True
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100


strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)


strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM prestamo_periodo ORDER BY id"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)


strCadena = "SELECT id_cuenta as Codigo, CONCAT(descripcion,'-',numero_cuenta,'    [',moneda,']') as Descripcion FROM view_cuenta_banco WHERE ruc='" & KEY_RUC & "' ORDER BY id_entidad"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentaBancaria)

strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)

strCadena = "SELECT dni as Codigo,nombre_completo  as Descripcion FROM view_entidad WHERE id_personal='si' and ruc='" & KEY_RUC & "' ORDER BY nombre_completo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPersonal)


strCadena = "SELECT dni as Codigo,nombre_completo  as Descripcion FROM view_entidad WHERE habilitado='si' and  id_cargo='00004' and  id_personal='si' and ruc='" & KEY_RUC & "' ORDER BY nombre_completo  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAutorizado)







Call llenarGrid(Me.HfgDetalle)


End Sub
Public Sub nuevo()
Me.TxtDetalle.Text = ""
Me.txtMonto.Text = ""
Me.TxtOperacion.Text = ""
Me.frmdetalle.Tag = 0
Me.frmdetalle.Visible = True
Me.cmdRegistrar.Enabled = True
Me.cmdTransferir.Enabled = False
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
Dim in_precio As String
On Error GoTo salir

strCadena = "call ADM_prestamo('3','','','','','','','','','','','','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 3500
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 2500
           Grilla.ColWidth(10) = 1500
    
         cabecera = "CODIGO" & vbTab & "FECHA PRESTAMO" & vbTab & "DNI" & vbTab & "PERSONA" & vbTab & "MONEDA" & vbTab & "MONTO" & vbTab & "SALDO" & vbTab & "CUENTA ORIGEN" & vbTab & "OPERACION" & vbTab & "AUTORIZADO POR" & vbTab & "PROCESADO"
         Grilla.AddItem cabecera
         For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        in_prestamo = 0
        in_saldo = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & rst("dni") & vbTab & rst("nombre_completo") & vbTab & rst("moneda") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00") & vbTab & rst("cuenta") & vbTab & rst("operacion") & vbTab & rst("autorizado") & vbTab & rst("estado")
            Grilla.AddItem Fila
            Grilla.col = 10
            Grilla.Row = i + 1
            
            If rst("estado") = "PENDIENTE" Then
                Grilla.CellBackColor = &H8080FF
            Else
                Grilla.CellBackColor = &H80FF80
            End If
            in_prestamo = in_prestamo + rst("monto")
            in_saldo = in_saldo + rst("saldo")
        
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_prestamo, "#,##0.00") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
        Grilla.AddItem Fila
        For k = 5 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

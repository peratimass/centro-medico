VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmCambioAceite 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   18000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCliente 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3960
      TabIndex        =   41
      Top             =   120
      Width           =   2535
   End
   Begin VitekeySoft.ChameleonBtn mdBuscar 
      Height          =   345
      Left            =   12960
      TabIndex        =   40
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "BUSCAR"
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
      MICON           =   "FrmCambioAceite.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpfechaInicio 
      Height          =   300
      Left            =   9840
      TabIndex        =   38
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
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
      Format          =   52166657
      CurrentDate     =   43495
   End
   Begin VB.TextBox txtPlacaBusqueda 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE DE CAMBIO DE ACEITE"
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
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   16695
      Begin VB.TextBox txtdni 
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
         Left            =   2040
         TabIndex        =   31
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SERVICIO TECNICO"
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
         Height          =   6855
         Left            =   6360
         TabIndex        =   19
         Top             =   360
         Width           =   10095
         Begin VB.TextBox txtCosto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   2160
            TabIndex        =   58
            Top             =   5520
            Width           =   1575
         End
         Begin VB.TextBox txtPrecio 
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
            Left            =   7920
            TabIndex        =   52
            Top             =   4920
            Width           =   1095
         End
         Begin VB.TextBox txtCantidad 
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
            Left            =   7080
            TabIndex        =   51
            Top             =   4920
            Width           =   735
         End
         Begin VB.TextBox txtDescripcionRepuesto 
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
            Left            =   3360
            TabIndex        =   50
            Top             =   4920
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   2280
            TabIndex        =   49
            Top             =   4920
            Width           =   975
         End
         Begin VB.TextBox txtKilometraje_final 
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
            Left            =   6120
            TabIndex        =   35
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtKilometraje_inicio 
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
            Left            =   2160
            TabIndex        =   34
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtobservacion 
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
            Height          =   915
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   1320
            Width           =   5655
         End
         Begin MSComCtl2.DTPicker DtpFecha 
            Height          =   345
            Left            =   2160
            TabIndex        =   21
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
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
            Format          =   52166657
            CurrentDate     =   43349
         End
         Begin MSDataListLib.DataCombo DtcTecnico 
            Height          =   315
            Left            =   2160
            TabIndex        =   25
            Top             =   2280
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4194304
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdsalir_detalle 
            Height          =   615
            Left            =   7440
            TabIndex        =   26
            Top             =   6000
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1085
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCambioAceite.frx":001C
            PICN            =   "FrmCambioAceite.frx":0038
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar 
            Height          =   615
            Left            =   3960
            TabIndex        =   27
            Top             =   6000
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            BTYPE           =   5
            TX              =   "GUARDAR"
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
            MICON           =   "FrmCambioAceite.frx":305F
            PICN            =   "FrmCambioAceite.frx":307B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker dtpProximoCambio 
            Height          =   320
            Left            =   2160
            TabIndex        =   30
            Top             =   2640
            Width           =   1815
            _ExtentX        =   3201
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
            Format          =   52166657
            CurrentDate     =   43349
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfdetalle 
            Height          =   1215
            Left            =   2160
            TabIndex        =   46
            Top             =   3360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2143
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   -2147483635
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
         Begin VitekeySoft.ChameleonBtn cmdImprimir 
            Height          =   615
            Left            =   5640
            TabIndex        =   47
            Top             =   6000
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1085
            BTYPE           =   5
            TX              =   "IMPRIMIR"
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
            MICON           =   "FrmCambioAceite.frx":66C3
            PICN            =   "FrmCambioAceite.frx":66DF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdQuitar 
            Height          =   615
            Left            =   9240
            TabIndex        =   48
            Top             =   3360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   7
            TX              =   ""
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCambioAceite.frx":8CB0
            PICN            =   "FrmCambioAceite.frx":8CCC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LISTADO DE REPUESTOS"
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
            Left            =   2160
            TabIndex        =   59
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO SERVICIO :"
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
            Left            =   600
            TabIndex        =   57
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
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
            Left            =   7080
            TabIndex        =   56
            Top             =   4680
            Width           =   705
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO"
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
            Left            =   8520
            TabIndex        =   55
            Top             =   4680
            Width           =   495
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION REPUESTO"
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
            Left            =   3360
            TabIndex        =   54
            Top             =   4680
            Width           =   1635
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CODIGO"
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
            Left            =   2280
            TabIndex        =   53
            Top             =   4680
            Width           =   555
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            BorderWidth     =   2
            Height          =   645
            Left            =   2160
            Top             =   4680
            Width           =   6975
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KM PROXIMO CAMBIO:"
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
            Left            =   4440
            TabIndex        =   45
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KM SERVICIO :"
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
            Left            =   1005
            TabIndex        =   33
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PROXIMO SERVICIO:"
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
            Left            =   600
            TabIndex        =   29
            Top             =   2760
            Width           =   1350
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TECNICO :"
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
            Left            =   1290
            TabIndex        =   24
            Top             =   2280
            Width           =   660
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OBSERVACION :"
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
            Left            =   915
            TabIndex        =   22
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA SERVICIO :"
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
            Left            =   795
            TabIndex        =   20
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.TextBox TxtMarca 
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
         Left            =   2040
         TabIndex        =   15
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtVim 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox TxtMotor 
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
         Left            =   2040
         TabIndex        =   12
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox TxtPlaca 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VitekeySoft.ChameleonBtn cmdSavePeriodo 
         Height          =   615
         Left            =   3960
         TabIndex        =   6
         Top             =   4920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "SAVE"
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
         MICON           =   "FrmCambioAceite.frx":B116
         PICN            =   "FrmCambioAceite.frx":B132
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdConsultar 
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "CONSULTAR"
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
         MICON           =   "FrmCambioAceite.frx":E77A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   6495
         Left            =   240
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblid_marca 
         Caption         =   "Label16"
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   6120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblid_cambio 
         Caption         =   "Label16"
         Height          =   255
         Left            =   840
         TabIndex        =   43
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblpropietario 
         BackColor       =   &H00404040&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   2040
         TabIndex        =   32
         Top             =   4200
         Width           =   3345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROPIETARIO DNI/ RUC :"
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
         Left            =   360
         TabIndex        =   17
         Top             =   3840
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Left            =   1410
         TabIndex        =   16
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° VIM :"
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
         Left            =   1455
         TabIndex        =   14
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° MOTOR  :"
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
         Left            =   1200
         TabIndex        =   11
         Top             =   2400
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° SERIE  :"
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
         Left            =   1365
         TabIndex        =   9
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLACA DE RODAJE :"
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
         Left            =   345
         TabIndex        =   7
         Top             =   840
         Width           =   1275
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   16920
      TabIndex        =   1
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCambioAceite.frx":E796
      PICN            =   "FrmCambioAceite.frx":E7B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfmensualidad 
      Height          =   7335
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   12938
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
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
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   855
      Left            =   16920
      TabIndex        =   3
      Top             =   2325
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ANULAR"
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
      MICON           =   "FrmCambioAceite.frx":EC04
      PICN            =   "FrmCambioAceite.frx":EC20
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
      Height          =   855
      Left            =   16920
      TabIndex        =   4
      Top             =   4070
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCambioAceite.frx":1106A
      PICN            =   "FrmCambioAceite.frx":11086
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdUpdate 
      Height          =   855
      Left            =   16920
      TabIndex        =   36
      Top             =   1455
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "VISUALIZAR"
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
      MICON           =   "FrmCambioAceite.frx":11476
      PICN            =   "FrmCambioAceite.frx":11492
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   300
      Left            =   11400
      TabIndex        =   39
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
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
      Format          =   52166657
      CurrentDate     =   43495
   End
   Begin VitekeySoft.ChameleonBtn cmdHistorial 
      Height          =   855
      Left            =   16920
      TabIndex        =   60
      Top             =   3200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "HISTORIAL"
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
      MICON           =   "FrmCambioAceite.frx":13ACB
      PICN            =   "FrmCambioAceite.frx":13AE7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE:"
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
      Left            =   3240
      TabIndex        =   42
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHAS:"
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
      Left            =   9120
      TabIndex        =   37
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° PLACA :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8040
      Left            =   0
      Top             =   0
      Width           =   18000
   End
End
Attribute VB_Name = "FrmCambioAceite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChameleonBtn2_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdConsultar_Click()
Call get_datos_placa(Me.TxtPlaca.Text)
End Sub

Private Sub get_cambio_aceite(ByVal in_cambio As String)
strCadena = "SELECT * FROM cambio_aceite WHERE id_cambio='" & Val(in_cambio) & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   Me.lblid_cambio.Caption = in_cambio
   Call get_datos_placa(rstA("placa"))
   Me.TxtPlaca.Text = rstA("placa")
   Me.DtpFecha.Value = rstA("fecha_emision")
   txtKilometraje_inicio.Text = rstA("kilometraje_cambio")
   txtKilometraje_final.Text = rstA("kilometraje_proximo_cambio")
   txtobservacion.Text = rstA("observacion")
   DtcTecnico.BoundText = rstA("id_tecnico")
   dtpProximoCambio.Value = rstA("fecha_proximo_cambio")
   Me.txtCosto.Text = rstA("costo_servicio")
   Me.frmdetalle.Visible = True
   
   Call Me.detalle(Me.Hfdetalle)
End If
End Sub

Private Sub cmdEliminar_Click()
Procedencia = anular
Call disabled_form(Me)
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdHistorial_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
Dim in_total As String

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = KEY_RUC
arr(1, 2) = KEY_EMPRESA


param = arr()

  strCadena = "call put_cambio_aceite_v2('8','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.txtPlacaBusqueda.Text) & "','00','0.00','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','0','0','" & KEY_RUC & "')"
  Call ConfiguraRst(strCadena)
  

  
  Ans = ShowMultiReport(rst, "rpt_servicio_tecnicoHistorial", param, App.Path + "\Reportes\")
  
   
End Sub

Private Sub cmdImprimir_Click()
  Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
Dim in_total As String

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = KEY_RUC
arr(1, 2) = KEY_EMPRESA


param = arr()

  strCadena = "call put_cambio_aceite_v2('7','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','00','0.00','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','0','0','" & KEY_RUC & "')"
  Call ConfiguraRst(strCadena)
  
  strCadena = "call put_cambio_aceite_v2('4','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','00','0.00','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','0','0','" & KEY_RUC & "')"
  Call ConfiguraRstK(strCadena)
  
  Ans = ShowMultiReport(rst, "rpt_servicio_tecnico", param, App.Path + "\Reportes\", , , True, , rstK, "rptservicio_tecnico_detalle")
  
   
  
End Sub

Private Sub cmdNuevo_Click()



Me.frmdetalle.Visible = True
Me.TxtPlaca.Text = ""
Me.TxtMarca.Text = ""
Me.txtdni.Text = ""
Me.txtVim.Text = ""
Me.TxtMotor.Text = ""
Me.txtSerie.Text = ""
Me.lblid_cambio.Caption = 0
Me.lblpropietario.Caption = ""
Call Resalta(Me.TxtPlaca)

End Sub

Private Sub cmdProcesar_Click()
If Trim(Me.txtdni.Text) <> "" Then
    
    strCadena = "call put_cambio_aceite_v2('1','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','00','0.00','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','0','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    MsgBox "Registro Exitoso.", vbInformation
   
   strCadena = "SELECT * FROM view_cambio_aceite WHERE   ruc='" & KEY_RUC & "'"
   Call listado(Me.hfmensualidad)
   Me.frmdetalle.Visible = False
   
    
End If
End Sub
Public Sub put_factura(ByVal in_propietario As String)
Call FrmVentas.activar
strCadena = "P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
in_dni = Trim(in_propietario)

FrmVentas.TxtCodCliente.Text = in_dni
Call FrmVentas.precionar_cliente

FrmVentas.txtobservacion.Text = "PROX CAMBIO:" & Space(1) & Format(Me.dtpProximoCambio.Value, "dd-mm-YYYY") & Space(2) & Trim(Me.txtKilometraje_final.Text) & "KM"

'strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
"('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','" & Trim(Me.txtid_servicio.Text) & "','1'," & _
"'" & Val(Me.lblprecio.Caption) & "','" & Val(Me.lblprecio.Caption) & "','0','si','" & Trim(Me.lblservicio.Caption) & "','" & KEY_USUARIO & "')"
'CnBd.Execute (strCadena)

'Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))
       
End Sub

Private Sub CmdQuitar_Click()

 strCadena = "call put_cambio_aceite_v2('6','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.txtCodigo.Text) & "','" & Val(Me.txtPrecio.Text) & "','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(Me.Hfdetalle.TextMatrix(Me.Hfdetalle.Row, 0)) & "','" & KEY_RUC & "')"
 CnBd.Execute (strCadena)
 Call Me.detalle(Me.Hfdetalle)


End Sub

Private Sub cmdSalir_Click()
Unload Me

End Sub

Private Sub cmdSalirPeriodo_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdsalir_detalle_Click()




Me.frmdetalle.Visible = False
End Sub

Private Sub cmdSavePeriodo_Click()
If Trim(Me.TxtPlaca.Text) <> "" And Trim(Me.TxtMarca.Text) <> "" Then
  
    strCadena = "call sp_marca_placa_v3('" & Val(Me.lblid_marca.Caption) & "','" & txtdni.Text & "','" & Trim(Me.TxtMarca.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','-','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtMotor.Text) & "','" & Trim(Me.txtVim.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    
    
    MsgBox "Registro Correcto", vbInformation

End If
End Sub

Private Sub cmdupdate_Click()
If Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0)) > 0 Then
    Call get_cambio_aceite(Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0)))
End If
End Sub



Private Sub Form_Load()
CenterForm Me
Me.Top = 100
DtpFecha.Value = KEY_FECHA
Me.DtpfechaInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTecnico)

dtpProximoCambio.Value = DateAdd("d", 30, Me.DtpFecha.Value)

strCadena = "SELECT * FROM view_cambio_aceite WHERE ruc='" & KEY_RUC & "' ORDER BY id_cambio DESC LIMIT 27"
Call listado(Me.hfmensualidad)
End Sub



Public Sub listado(ByVal Grilla As MSHFlexGrid)
'On Error GoTo SALIR
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
       
       Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 900
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 3300
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 2500
           Grilla.ColWidth(7) = 4000
        Next
         cabecera = "CODIGO" & vbTab & "FECHA " & vbTab & "PLACA" & vbTab & "DNI/RUC" & vbTab & "PROPIETARIO" & vbTab & "COSTO SERVICIO" & vbTab & "TECNICO" & vbTab & "OBSERVACION"
         Grilla.AddItem cabecera
         For k = 0 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = Format(rst("id_cambio"), "00000") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("placa") & vbTab & rst("dni") & vbTab & rst("propietario") & vbTab & rst("precio") & vbTab & rst("tecnico") & vbTab & rst("observacion")
             Grilla.AddItem Fila
             If rst("anulado") = "si" Then
                            For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
        
             End If
        rst.MoveNext
        Next i
        
Exit Sub
'SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Public Sub detalle(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

If Val(Me.lblid_cambio.Caption) > 0 Then
    in_operacion = 4
Else
    in_operacion = 5
End If

strCadena = "call put_cambio_aceite_v2('" & in_operacion & "','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.txtCodigo.Text) & "','" & Val(Me.txtPrecio.Text) & "','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','" & Val(Me.txtCantidad.Text) & "','0','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
       
       Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 3200
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 900
           Grilla.ColWidth(5) = 900
          
        Next
         cabecera = "ID" & vbTab & "CODIGO" & vbTab & "REPUESTO " & vbTab & "CANT" & vbTab & "PRECIO" & vbTab & "TOTAL"
         Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        in_total = 0
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("cantidad") & vbTab & rst("precio") & vbTab & rst("total")
             Grilla.AddItem Fila
             in_total = in_total + rst("total")
             rst.MoveNext
        Next i
         Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_total, "#,##0.00")
         Grilla.AddItem Fila
         
         For k = 2 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
          Next k
          
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub




Private Sub get_datos_placa(ByVal in_placa As String)
If Len(in_placa) > 3 Then
strCadena = "SELECT * FROM persona_transporte WHERE placa='" & Trim(in_placa) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblid_marca.Caption = rst("id")
    Me.TxtPlaca.Text = in_placa
    Me.txtVim.Text = rst("vim")
    Me.TxtMarca.Text = rst("marca")
    Me.TxtMotor.Text = rst("motor")
    Me.txtSerie.Text = rst("serie")
    Me.txtdni.Text = rst("id_persona")
    Me.lblpropietario = get_persona(Trim(Me.txtdni.Text))
Else
    Me.lblid_marca.Caption = ""
    Me.txtVim.Text = ""
    Me.TxtMarca.Text = ""
    Me.TxtMotor.Text = ""
    Me.txtSerie.Text = ""
    Me.txtdni.Text = ""
    Me.lblpropietario.Caption = ""
    
        
End If
Else
    MsgBox "Ingrese una PLACA CORRECTA", vbInformation
End If

End Sub



Public Sub get_propietariov2(ByVal in_dni As String)
buscar_nuevamente:
If Len(in_dni) >= 8 Then
strCadena = "SELECT * FROM  view_entidad WHERE  dni='" & Trim(in_dni) & "'LIMIT 1 "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        If Len(Trim(in_dni)) = 8 Then
            If get_dni_reniec_ii(Trim(in_dni)) = True Then
                GoTo buscar_nuevamente
            End If
        End If
        
        Procedencia = 1
        FrmDetallePersona.Show
        If Len(Trim(in_dni)) = 8 Then
            nruc = "10" & Trim(in_dni)
            FrmDetallePersona.txtRuc.Text = DigitoVerificadorRUC(Trim(nruc))
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        Else
            FrmDetallePersona.txtRuc.Text = Trim(in_dni)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        End If
    Else
        Me.lblpropietario.Caption = rst("nombre_completo")
    End If
Else
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
End Sub

Private Sub mdBuscar_Click()
strCadena = "SELECT * FROM view_cambio_aceite WHERE fecha_emision>='" & Format(Me.DtpfechaInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' ORDER BY id_cambio DESC "
    Call listado(Me.hfmensualidad)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call Resalta(Me.txtPrecio)
End If


End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_cambio_aceite WHERE propietario LIKE '%" & Trim(Me.txtCliente.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY id_cambio DESC"
    Call listado(Me.hfmensualidad)
End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If

End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call get_propietariov2(Me.txtdni.Text)
End If
End Sub

Private Sub txtid_servicio_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txtKilometraje_inicio_Change()
txtKilometraje_final.Text = Val(Me.txtKilometraje_inicio.Text) + 4000
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call get_datos_placa(Me.TxtPlaca.Text)
End If
End Sub

Private Sub txtPlacaBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_cambio_aceite WHERE placa LIKE '%" & Trim(Me.txtPlacaBusqueda.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call listado(Me.hfmensualidad)
End If
End Sub

Private Sub put_detalle()

 strCadena = "call put_cambio_aceite_v2('2','" & Val(Me.lblid_cambio.Caption) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtdni.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.txtCodigo.Text) & "','" & Val(Me.txtPrecio.Text) & "','" & Me.DtcTecnico.BoundText & "','" & Trim(Me.txtobservacion.Text) & "','" & Val(Me.txtKilometraje_inicio.Text) & "','" & Val(Me.txtKilometraje_final.Text) & "','" & Format(Me.dtpProximoCambio.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & Val(Me.txtCosto.Text) & "','" & Val(Me.txtCantidad.Text) & "','0','" & KEY_RUC & "')"
 CnBd.Execute (strCadena)
 Call Me.detalle(Me.Hfdetalle)
 


End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call put_detalle
   Me.txtCodigo.Text = ""
   Me.txtDescripcionRepuesto.Text = ""
   Me.txtCantidad.Text = ""
   Me.txtPrecio.Text = 0
   Call Resalta(Me.txtCodigo)
End If
End Sub

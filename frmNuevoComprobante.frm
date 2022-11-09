VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmNuevoComprobante 
   BorderStyle     =   0  'None
   Caption         =   "INGRESO DE DOCUMENTOS"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Comprobante"
      ForeColor       =   &H00800000&
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtEfectivo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11400
         TabIndex        =   78
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtCuentaCorriente 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   11400
         TabIndex        =   77
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ver"
         Height          =   255
         Left            =   14160
         TabIndex        =   76
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ver"
         Height          =   255
         Left            =   14160
         TabIndex        =   75
         Top             =   1440
         Width           =   615
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4335
         Left            =   5400
         TabIndex        =   42
         Top             =   2760
         Width           =   9495
         Begin VB.Frame FrmCheque 
            Caption         =   "PAGAR CON CHEQUE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1575
            Left            =   360
            TabIndex        =   66
            Top             =   2400
            Width           =   4335
            Begin VB.CommandButton cmdCargarCheque 
               Caption         =   "Cargar Cheque"
               Height          =   255
               Left            =   1680
               TabIndex        =   69
               Top             =   1200
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.OptionButton OptChequeNO 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "No"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   240
               TabIndex        =   68
               Top             =   360
               Width           =   735
            End
            Begin VB.OptionButton OptChequeSi 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Si"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   960
               TabIndex        =   67
               Top             =   360
               Width           =   615
            End
            Begin MSDataListLib.DataCombo DtcCheque 
               Height          =   315
               Left            =   1680
               TabIndex        =   70
               Top             =   360
               Visible         =   0   'False
               Width           =   2535
               _ExtentX        =   4471
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
            Begin VB.Image Image1 
               Height          =   255
               Left            =   240
               Top             =   1200
               Width           =   1335
            End
         End
         Begin VB.Frame FrmaeMonto 
            Height          =   1335
            Left            =   360
            TabIndex        =   59
            Top             =   960
            Width           =   4335
            Begin VB.TextBox TxtMonto1 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   720
               MaxLength       =   80
               TabIndex        =   60
               Top             =   480
               Width           =   855
            End
            Begin MSDataListLib.DataCombo txtOrigen 
               Height          =   345
               Left            =   1680
               TabIndex        =   61
               Top             =   480
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   609
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               ForeColor       =   8388608
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "ORIGEN.PAGO"
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
               Left            =   1920
               TabIndex        =   65
               Top             =   120
               Width           =   1140
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "IMPORTES"
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
               Left            =   720
               TabIndex        =   64
               Top             =   120
               Width           =   870
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Monto:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   63
               Top             =   480
               Width           =   495
            End
            Begin VB.Label lblDescripcion1 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   1680
               TabIndex        =   62
               Top             =   840
               Width           =   2040
            End
         End
         Begin VB.Frame FrameCCostos 
            Height          =   3015
            Left            =   4800
            TabIndex        =   43
            Top             =   960
            Width           =   4455
            Begin VB.TextBox Monto3 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2880
               MaxLength       =   80
               TabIndex        =   53
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox TxtNaturaleza 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   120
               MaxLength       =   80
               TabIndex        =   52
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtCostos1 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   51
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtCostos3 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   50
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox txtCostos2 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   49
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox txtCostos4 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   48
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox Monto1 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2880
               MaxLength       =   80
               TabIndex        =   47
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox Monto2 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2880
               MaxLength       =   80
               TabIndex        =   46
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox Monto4 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2880
               MaxLength       =   80
               TabIndex        =   45
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox TxtTotalCC 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00800000&
               Height          =   285
               Left            =   2880
               MaxLength       =   80
               TabIndex        =   44
               Top             =   2040
               Width           =   1095
            End
            Begin VB.Label LblDescripcion2 
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   1275
               Left            =   120
               TabIndex        =   58
               Top             =   960
               Width           =   1320
            End
            Begin VB.Label Label14 
               Caption         =   "Gº. NATURALEZA"
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
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label18 
               Caption         =   "C.COSTOS"
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
               Left            =   1920
               TabIndex        =   56
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "IMPORTE"
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
               Left            =   3360
               TabIndex        =   55
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "TOTAL C.C"
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
               Left            =   1560
               TabIndex        =   54
               Top             =   2040
               Width           =   840
            End
         End
         Begin MSDataListLib.DataCombo DtcIgv 
            Height          =   315
            Left            =   1560
            TabIndex        =   71
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   "DataCombo1"
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H80000005&
            BorderStyle     =   3  'Dot
            FillColor       =   &H000000C0&
            Height          =   3270
            Left            =   120
            Top             =   840
            Width           =   9255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Afecto IGV:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   480
            TabIndex        =   72
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame FrameComprobante 
         Caption         =   "COMPROBANTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Left            =   360
         TabIndex        =   33
         Top             =   1080
         Width           =   4695
         Begin VB.TextBox TxtNumero 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            MaxLength       =   80
            TabIndex        =   40
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtSerie 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1200
            MaxLength       =   80
            TabIndex        =   39
            Top             =   240
            Width           =   870
         End
         Begin VB.TextBox TxtTD 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   120
            MaxLength       =   80
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox TxtTotalImporte 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            MaxLength       =   80
            TabIndex        =   37
            Top             =   600
            Width           =   1935
         End
         Begin VB.Frame FrameIntacta 
            Height          =   550
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   4215
            Begin VB.TextBox TxtMontoFactura 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   315
               Left            =   1920
               MaxLength       =   80
               TabIndex        =   35
               Top             =   160
               Width           =   1935
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Importe Factura:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   360
               TabIndex        =   36
               Top             =   160
               Width           =   1380
            End
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Importe Saldo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   720
            TabIndex        =   41
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ENTIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   360
         TabIndex        =   29
         Top             =   4440
         Width           =   4695
         Begin VB.TextBox txtCodPersona 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            MaxLength       =   80
            TabIndex        =   83
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtRuc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   120
            MaxLength       =   80
            TabIndex        =   30
            Top             =   360
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo DtcEntidad 
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ruc"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1800
            TabIndex        =   32
            Top             =   360
            Width           =   300
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cuentas x Cobrar"
         Height          =   1095
         Left            =   7920
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "TIPO MOVIMIENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   4695
         Begin MSDataListLib.DataCombo DtcMovimiento 
            Height          =   315
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   "DataCombo1"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Cambio"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   5280
         TabIndex        =   24
         Top             =   1680
         Width           =   2535
         Begin VB.TextBox TxtTC 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1200
            MaxLength       =   80
            TabIndex        =   25
            Top             =   200
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1335
         Left            =   5280
         TabIndex        =   19
         Top             =   240
         Width           =   2535
         Begin MSComCtl2.DTPicker dtpValor 
            Height          =   300
            Left            =   840
            TabIndex        =   20
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            Format          =   121176065
            CurrentDate     =   40610
         End
         Begin MSComCtl2.DTPicker DtpPago 
            Height          =   300
            Left            =   840
            TabIndex        =   21
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            Format          =   121176065
            CurrentDate     =   40610
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Emision:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Pago:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   750
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "TOTAL MOVIMIENTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   360
         TabIndex        =   13
         Top             =   2880
         Width           =   4695
         Begin VB.TextBox TxtOperacion 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   15
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox TxtIdCompra 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   14
            Top             =   960
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DtcMoneda 
            Height          =   315
            Left            =   1920
            TabIndex        =   16
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº Operacion:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   480
            TabIndex        =   18
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   840
            TabIndex        =   17
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "GLOSA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   360
         TabIndex        =   11
         Top             =   5640
         Width           =   4695
         Begin VB.TextBox TxtGlosa 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   795
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ctas PagarProveedores"
         Height          =   495
         Left            =   9240
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdListado 
         Caption         =   "Recibos de Egreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   9
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Ordenes de Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   8
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ingresos Efectuados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   7
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ctas Pagar Otros"
         Height          =   615
         Left            =   9240
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame FrameFacturas 
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
         Height          =   1695
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox TxtTotalFacturas 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2520
            MaxLength       =   80
            TabIndex        =   3
            Top             =   1320
            Width           =   1935
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
            Height          =   1215
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   2143
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8388608
            FixedCols       =   0
            ForeColorFixed  =   8388608
            GridColor       =   0
            FocusRect       =   0
            GridLines       =   2
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Total Importe:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   840
            TabIndex        =   5
            Top             =   1320
            Width           =   1200
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Gastos sin Destino"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   1
         Top             =   2400
         Width           =   3255
      End
      Begin ComCtl3.CoolBar ClbAcciones 
         Height          =   2490
         Left            =   9000
         TabIndex        =   73
         Top             =   7320
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   4392
         BandCount       =   1
         ForeColor       =   -2147483635
         FixedOrder      =   -1  'True
         VariantHeight   =   0   'False
         EmbossPicture   =   -1  'True
         _CBWidth        =   5835
         _CBHeight       =   2490
         _Version        =   "6.0.8169"
         Child1          =   "TlbAcciones"
         MinHeight1      =   2430
         Width1          =   3180
         FixedBackground1=   0   'False
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar TlbAcciones 
            Height          =   810
            Left            =   30
            TabIndex        =   74
            Top             =   30
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   1429
            ButtonWidth     =   1852
            ButtonHeight    =   1429
            Style           =   1
            ImageList       =   "ImgIconos"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nuevo"
                  Key             =   "(Nuevo)"
                  Object.ToolTipText     =   "Grabar Ctrl+G"
                  ImageKey        =   "(Nuevo)"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Grabar"
                  Key             =   "(Grabar)"
                  ImageKey        =   "(Grabar)"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Rbo Egreso"
                  Key             =   "(Imprimir)"
                  ImageKey        =   "(Imprimir)"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Orden Pago"
                  Key             =   "(OrdenPago)"
                  ImageKey        =   "(Imprimir)"
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Cancelar"
                  Key             =   "(Salir)"
                  Object.ToolTipText     =   "Cancelar"
                  ImageKey        =   "(Salir)"
               EndProperty
            EndProperty
            OLEDropMode     =   1
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
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNuevoComprobante.frx":0000
               Key             =   "(Modificar)"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNuevoComprobante.frx":031C
               Key             =   "(Imprimir)"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNuevoComprobante.frx":03A9
               Key             =   "(Nuevo)"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNuevoComprobante.frx":0809
               Key             =   "(Salir)"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmNuevoComprobante.frx":0C69
               Key             =   "(Grabar)"
            EndProperty
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVinculados 
         Height          =   1455
         Left            =   360
         TabIndex        =   82
         Top             =   6960
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2566
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         GridColor       =   0
         FocusRect       =   0
         GridLines       =   2
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
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "SALDOS CUENTAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   12240
         TabIndex        =   81
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CTA CORRIENT"
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
         Height          =   255
         Left            =   12960
         TabIndex        =   80
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "EFECTIVO"
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
         Height          =   255
         Left            =   12960
         TabIndex        =   79
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000C0&
         Height          =   1710
         Left            =   11280
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmNuevoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo As String
Dim escaneo As String
Public Procedencia As EnumProcede
Dim img As String
Dim moneda As String
Dim fecha_valor As Date
Dim fecha_sis As Date
Dim Monto As Double
Dim t_cambio As Single
Dim cuenta As String
Dim Saldo As Double
Dim Num_registros As Integer
Dim codigo_P As String
Public busqueda As EnumCostos
Dim idRecibo As Double
Dim nordenpago As Double

Private Sub cmdcomprobantesrel_Click()
FrmListadoFacturasVarios.Show
End Sub
Public Sub LlenarVinculados(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT     DocumentoCompra.idCompra, DocumentoCompra.dVencimiento,(Comprobantes.doc_abrev +':'+ DocumentoCompra.sSerie +'-'+ " & _
"DocumentoCompra.cDocumentoCompra) as Numero , DocumentoCompra.Saldo,seleccion FROM DocumentoCompra INNER JOIN " & _
"Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE cPersona='" & cPersona & "' AND Ruc='" & KEY_RUC & "' AND saldo>0 AND Anulado='F'"
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
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2400
           Grilla.ColWidth(3) = 1200
         Next
        cabecera = "IdCompra" & vbTab & "Vencimiento" & vbTab & "Documento" & vbTab & "Saldo"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("idCompra") & vbTab & rst("dVencimiento") & vbTab & rst("Numero") & vbTab & Format(rst("Saldo"), "#,##0.00")
            Grilla.AddItem Fila
            If rst("seleccion") = "V" Then
                For k = 0 To 3
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
            End If
            tTotal = tTotal + rst("saldo")
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      
            Grilla.col = 3
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
 
 
Private Sub CmdFoto_Click()
'Me.CommonDialog1.Filter = "*.jpg"
'Me.CommonDialog1.ShowOpen
'Me.ImgEscaneo.Picture = LoadPicture(Me.CommonDialog1.FileName)
'img = Me.CommonDialog1.FileName
End Sub

Private Sub DtcCCostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtGlosa)
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub cmdListado_Click()
FrmRecibosEgresoLIst.Show
End Sub

Private Sub Command1_Click()
FrmCuentasxCobrar.Show
End Sub

Private Sub Command2_Click()
'strCadena = "UPDATE DocumentoCompra SET seleccion='no' "
'CnBd.Execute (strCadena)
Procedencia = buscar
FrmListadoFacturasCompra.Show
End Sub
Public Sub llenarFacturas()
strCadena = "SELECT dEmisionCompra as EMISION,(Comprobantes.doc_abrev +':'+ DocumentoCompra.sSerie +'-'+ DocumentoCompra.cDocumentoCompra)AS COMPROBANTE, DocumentoCompra.saldo as SALDO " & _
            "FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE seleccion='si' AND IdUsuario='" & Trim(KEY_USUARIO) & "' ORDER BY dEmisionCompra ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        Me.HfgFacturas.Rows = 1
    Set Me.HfgFacturas.Recordset = rst
        Me.HfgFacturas.Rows = rst.RecordCount
        Me.HfgFacturas.ColWidth(0) = 970
        Me.HfgFacturas.ColWidth(1) = 2500
        Me.HfgFacturas.ColWidth(2) = 800
End If
Set rst = Nothing
End Sub


Private Sub Command4_Click()
strCadena = "UPDATE DocumentoCompra SET seleccion='no' "
CnBd.Execute (strCadena)
 
Procedencia = buscar
FrmListadoFacturasVarios.Show
End Sub

Private Sub Command3_Click()
FrmReciboIngresosList.Show
End Sub

Private Sub Command5_Click()
FrmMiscuentas.Show
End Sub

Private Sub Command6_Click()
FrmMiscuentas.Show
End Sub

Private Sub Command7_Click()
FrmOrdenPagoList.Show
End Sub

Private Sub Command8_Click()
Procedencia = buscar
FrmListadoFacturasSinDestino.Show
Exit Sub
End Sub

Private Sub DtcIgv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim cPersona As String
strCadena = "SELECT * FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    cPersona = rst("cPersona")
End If
Set rst = Nothing

    If MsgBox("Desea pagar al Contado ?", vbQuestion + vbYesNo, "Mensaje para el Usuario") = vbYes Then
        
        Me.FrameCCostos.Visible = False
        If Me.DtcIgv.BoundText = "si" Then
            vimponible = Val(Me.TxtTotalImporte.Text) / (KEY_IGV + 1)
            vigv = Val(vimponible * KEY_IGV)
            
        Else
            vimponible = Val(Me.TxtTotalImporte.Text)
            vigv = 0#
        End If
        
        
        
        strCadena = "INSERT INTO DocumentoCompra(cDocumentoCompra,doc_cod,sSerie,cPersona,Persona,Observacion,Alm_cod,moneda,tc, " & _
            "dEmisionCompra,nSubTotal,nIgv,nTotalCompra,FechaProceso," & _
            "Anulado,IdUsuario,saldo,tipo_factura)VALUES ('" & Trim(Me.txtnumero.Text) & "','" & Trim(Me.TxtTD.Text) & "', " & _
            "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(cPersona) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','0001'," & _
            "'" & Trim(Me.DtcMoneda.BoundText) & "','" & Val(Me.TxtTc.Text) & "','" & CVDate(Me.DtpValor.Value) & "','" & Val(vimponible) & "'," & _
            "'" & Val(vigv) & "','" & Val(Me.TxtTotalImporte.Text) & "','" & CVDate(Date) & "','no','" & Trim(KEY_USUARIO) & "','" & Val(Me.TxtTotalImporte.Text) & "'," & _
            "'egreso')"
            CnBd.Execute (strCadena)
             
                 
           strCadena = "SELECT * FROM DocumentoCompra WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "' AND cPersona='" & Trim(cPersona) & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                Me.TxtIdCompra.Text = rst(0)
            Else
            MsgBox "Hubo un Error al Grabar el Comprobante", vbInformation, "Mensaje para el Usuario"
           Exit Sub
           End If
        
        Me.TxtMonto1.Text = Format(Val(Me.TxtTotalImporte.Text), "###0.00")
        Call Resalta(Me.TxtMonto1)
        Me.FrameCCostos.Visible = False
        
    Else
    
        Me.FrameCCostos.Visible = False
        strCadena = "INSERT INTO DocumentoCompra(cDocumentoCompra,doc_cod,sSerie,cPersona,Persona,Observacion,Alm_cod,moneda,tc, " & _
            "dEmisionCompra,nSubTotal,nIgv,nTotalCompra,FechaProceso," & _
            "Anulado,IdUsuario,saldo,tipo_factura)VALUES ('" & Trim(Me.txtnumero.Text) & "','" & Trim(Me.TxtTD.Text) & "', " & _
            "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(cPersona) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','0001'," & _
            "'" & Trim(Me.DtcMoneda.BoundText) & "','" & Val(Me.TxtTc.Text) & "','" & CVDate(Me.DtpValor.Value) & "','" & Val(vimponible) & "'," & _
            "'" & Val(vigv) & "','" & Val(Me.TxtTotalImporte.Text) & "','" & CVDate(Date) & "','no','" & Trim(KEY_USUARIO) & "','" & Val(Me.TxtTotalImporte.Text) & "'," & _
            "'egreso')"
            CnBd.Execute (strCadena)
             
                       strCadena = "SELECT * FROM DocumentoCompra WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "' AND cPersona='" & Trim(cPersona) & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                Me.TxtIdCompra.Text = rst(0)
            Else
            MsgBox "Hubo un Error al Grabar el Comprobante", vbInformation, "Mensaje para el Usuario"
           Exit Sub
           End If
    MsgBox "Comprobante Guardado Exitosamente!!"
    Call nuevo
    End If
    Call Resalta(Me.TxtMonto1)
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtOperacion)
End If
End Sub

Private Sub DtcMovimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtTD)
End If
End Sub

Private Sub Form_Activate()

    'Me.DtcMovimiento.SetFocus

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpPago.Value = CVDate(KEY_FECHA)
Me.DtpValor.Value = CVDate(KEY_FECHA)

Me.OptChequeNO.Value = True
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM Moneda ORDER BY id_moneda ASC "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
  Set rst = Nothing
  
  
  strCadena = "SELECT igv as Codigo, igv as Descripcion FROM igv ORDER BY igv ASC "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcIgv)
  Set rst = Nothing
  
  strCadena = "SELECT     mis_cuentas.id_cuenta as Codigo, plan_contable_det.plan_des as Descripcion FROM mis_cuentas INNER JOIN " & _
  "plan_contable_det ON mis_cuentas.cuenta_ctble = plan_contable_det.pc_codigo WHERE plan_contable_det.id_plancontable='0001'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.txtOrigen)
  Set rst = Nothing
  
  strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM tipo_movimiento ORDER BY id_tipo ASC "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMovimiento)
  Set rst = Nothing
  
  Me.TxtTc.Text = cambio(KEY_FECHA)
  fecha_sis = key_date
  Me.DtpValor.Value = CVDate(KEY_FECHA)
  Me.DtpPago.Value = CVDate(KEY_FECHA)
  Me.TlbAcciones.Buttons(KEY_ORDENPAGO).Enabled = False
  Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = False
  Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
  
Call ActualizarSaldos
End Sub

Public Sub ActualizarSaldos()
strCadena = "SELECT sum(mis_cuentas_det.montoreal) FROM mis_cuentas INNER JOIN mis_cuentas_det ON mis_cuentas.id_cuenta = mis_cuentas_det.id_cuenta WHERE mis_cuentas.tipo_cuenta='caja' AND fecha<='" & CVDate(KEY_FECHA) & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Me.TxtEfectivo.Text = Format(0, "#,##0.00")
Else
    Me.TxtEfectivo.Text = Format(rst(0), "#,##0.00")
End If
Set rst = Nothing

strCadena = "SELECT sum(mis_cuentas_det.montoreal) FROM mis_cuentas INNER JOIN mis_cuentas_det ON mis_cuentas.id_cuenta = mis_cuentas_det.id_cuenta WHERE mis_cuentas.tipo_cuenta='banco' AND fecha<='" & CVDate(KEY_FECHA) & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Me.TxtCuentaCorriente.Text = Format(0, "#,##0.00")
Else
    Me.TxtCuentaCorriente.Text = Format(rst(0), "#,##0.00")
End If
Set rst = Nothing

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
End If
End Sub

Private Sub HfVinculados_DblClick()
Dim cPersona As Double
If Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) > 0 Then
    strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 And Val(Me.TxtIdCompra.Text) <> Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) Then
        If rst("seleccion") = "F" Then
            strCadena = "UPDATE DocumentoCompra SET seleccion='V'WHERE idCompra='" & rst("idCompra") & "'"
        Else
            strCadena = "UPDATE DocumentoCompra SET seleccion='F'WHERE idCompra='" & rst("idCompra") & "'"
        End If
        cPersona = rst("cPersona")
        CnBd.Execute (strCadena)
         
        Call LlenarVinculados(Me.HfVinculados, cPersona)
        Saldo = 0
       strCadena = "SELECT Comprobantes.doc_abrev , DocumentoCompra.sSerie , DocumentoCompra.cDocumentoCompra AS Numero,DocumentoCompra.Saldo " & _
       "FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE DocumentoCompra.Ruc='" & KEY_RUC & "' AND saldo>0 AND anulado='F' AND cPersona='" & cPersona & "' AND seleccion='V' "
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
                If rst.RecordCount > 1 Then
                    barra = "/"
                Else
                    barra = ""
                End If
            For i = 0 To rst.RecordCount - 1
                
               Saldo = Saldo + rst("saldo")
               comprobante = Mid(rst("doc_abrev"), 1, 3) + ":" + Mid(rst("sSerie"), 2, 4) + "-" + Mid(rst("numero"), 5, 10) + barra + comprobante
               rst.MoveNext
            Next i
            Me.TxtGlosa.Text = "PAGO:" + comprobante
            Me.TxtMonto1.Text = Format(Saldo, "###0.00")
            Call Resalta(Me.TxtMonto1)
        End If
End If
End If
End Sub

Private Sub Monto1_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(Me.TxtMontoFactura.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto1)
End If
End Sub

Private Sub Monto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto1.Text = Format(Val(Me.Monto1.Text), "###0.00")
    Call Resalta(Me.txtCostos2)
End If
End Sub

Private Sub Monto2_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(Me.TxtMontoFactura.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto2)
End If
End Sub

Private Sub Monto2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto2.Text = Format(Val(Me.Monto2.Text), "###0.00")
    Call Resalta(Me.txtCostos3)
End If
End Sub

Private Sub Monto3_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(Me.TxtMontoFactura.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto3)
End If
End Sub

Private Sub Monto3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto3.Text = Format(Val(Me.Monto3.Text), "###0.00")
    Call Resalta(Me.txtCostos4)
End If
End Sub

Private Sub Monto4_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(Me.TxtMontoFactura.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto4)
End If
End Sub

Private Sub Monto4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto4.Text = Format(Val(Me.Monto4.Text), "###0.00")
     
End If
End Sub

Private Sub OptChequeNO_Click()
If Me.OptChequeNO.Value = True Then
    Me.DtcCheque.Visible = False
    Me.cmdCargarCheque.Visible = False
End If
End Sub

Private Sub OptChequeSi_Click()
Dim rstc As New ADODB.Recordset
If Me.OptChequeSi.Value = True Then
     strCadena = "SELECT     mis_cuentas.id_cuenta, mis_cuentas.descripcion " & _
        "FROM mis_cuentas  WHERE mis_cuentas.id_cuenta='" & Trim(Me.txtOrigen.BoundText) & "'"
        
        rstc.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    
     strCadena = "SELECT id_cheque as Codigo, id_cheque as Descripcion FROM Cheques WHERE id_cuenta='" & Val(rstc("id_cuenta")) & "' AND estado='libre'"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcCheque)
     Set rst = Nothing
     Me.DtcCheque.Visible = True
     Me.cmdCargarCheque.Visible = True
    Else
     Me.DtcCheque.Visible = False
     Me.cmdCargarCheque.Visible = False
    
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
        Call nuevo
    Case KEY_SAVE
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
       If Me.FrameCCostos.Visible = True And Me.TxtNaturaleza.Text <> "" Then
            Call Save_costosG
            Exit Sub
       End If
    
    If Me.FrameFacturas.Visible = True Then
            Call Save_varios
        Else
            Call Save
        End If
      Call ActualizarSaldos
      
      Case KEY_PRINT
         strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE codigo='" & idRecibo & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
        Exit Sub
    Case KEY_ORDENPAGO
                strCadena = "SELECT empresa, direccion, ruc_emp, serie, numero, num_cheque, entidad_financiera, cpersona, persona, fecha, cambio, glosa, monto, monto_letras " & _
                "FROM  OrdenPago WHERE idorden='" & nordenpago & "'"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptOrdenPago", , App.Path + "\Reportes\")
                Exit Sub
      Case KEY_EXIT
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub

End Sub
Private Sub Save_costosG()
If Me.TxtCostos1.Text <> "" Then
    strCadena = "INSERT INTO CentroCostosDoc (doc_cod,id_comprobante,cuenta_naturaleza,cuenta,monto)VALUES" & _
    "('" & Trim(Me.TxtTD.Text) & "','" & Val(Me.TxtIdCompra.Text) & "','" & Trim(Me.TxtNaturaleza.Text) & "','" & Trim(Me.TxtCostos1.Text) & "','" & Val(Me.Monto1.Text) & "')"
    CnBd.Execute (strCadena)
     
End If
If Me.txtCostos2.Text <> "" Then
    strCadena = "INSERT INTO CentroCostosDoc (doc_cod,id_comprobante,cuenta_naturaleza,cuenta,monto)VALUES" & _
    "('" & Trim(Me.TxtTD.Text) & "','" & Val(Me.TxtIdCompra.Text) & "','" & Trim(Me.TxtNaturaleza.Text) & "','" & Trim(Me.txtCostos2.Text) & "','" & Val(Me.Monto2.Text) & "')"
    CnBd.Execute (strCadena)
     
End If
If Me.txtCostos3.Text <> "" Then
    strCadena = "INSERT INTO CentroCostosDoc (doc_cod,id_comprobante,cuenta_naturaleza,cuenta,monto)VALUES" & _
    "('" & Trim(Me.TxtTD.Text) & "','" & Val(Me.TxtIdCompra.Text) & "','" & Trim(Me.TxtNaturaleza.Text) & "','" & Trim(Me.txtCostos3.Text) & "','" & Val(Me.Monto4.Text) & "')"
    CnBd.Execute (strCadena)
     
End If
If Me.txtCostos4.Text <> "" Then
    strCadena = "INSERT INTO CentroCostosDoc (doc_cod,id_comprobante,cuenta_naturaleza,cuenta,monto)VALUES" & _
    "('" & Trim(Me.TxtTD.Text) & "','" & Val(Me.TxtIdCompra.Text) & "','" & Trim(Me.TxtNaturaleza.Text) & "','" & Trim(Me.txtCostos4.Text) & "','" & Val(Me.Monto4.Text) & "')"
    CnBd.Execute (strCadena)
     
End If

If Trim(Me.TxtTD.Text) <> "0097" Then
    strCadena = "UPDATE DocumentoCompra SET destino='si' WHERE idCompra='" & Val(Me.TxtIdCompra.Text) & "'"
    CnBd.Execute (strCadena)
     
Else
    strCadena = "UPDATE movimiento_caja SET destino='si' WHERE codigo='" & Val(Me.TxtIdCompra.Text) & "'"
    CnBd.Execute (strCadena)
     
End If


Call nuevo

End Sub
Private Sub nuevo()
Me.FrameComprobante.Visible = True
Me.FrameFacturas.Visible = False
Me.TxtTD.Text = ""
Me.TxtSerie.Text = ""
Me.txtnumero.Text = ""
'Me.txtOrigen.Text = ""
Me.TxtNaturaleza.Text = ""
Me.TxtCostos1.Text = ""
Me.txtCostos2.Text = ""
Me.txtCostos3.Text = ""
Me.txtCostos4.Text = ""
Me.Monto1.Text = ""
Me.Monto2.Text = ""
Me.Monto3.Text = ""
Me.Monto4.Text = ""
Me.TxtTotalCC.Text = ""
Me.lblDescripcion1.Caption = ""
Me.LblDescripcion2.Caption = ""
Me.txtOperacion.Text = ""
Me.TxtTotalImporte.Text = 0#
Me.TxtMontoFactura.Text = 0#
Me.txtOperacion.Text = ""
Me.TxtIdCompra.Text = 0
Me.TxtRuc.Text = ""
Me.DtcEntidad.Text = ""
Me.TxtGlosa.Text = ""
Me.DtpValor.Value = KEY_FECHA
Me.DtpPago.Value = KEY_FECHA
Me.TxtMonto1.Text = ""
Me.OptChequeNO.Value = True
Me.FrmaeMonto.Visible = True
Me.FrmCheque.Visible = True
Me.FrameCCostos.Visible = True
Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
Me.DtcMovimiento.SetFocus
End Sub
Private Sub Save()
Dim Saldo As Single
Dim monto_letras As String
Dim Monto As Double
Dim glosaITF As String
Dim rstc As New ADODB.Recordset
On Error GoTo salir
Monto = Val(Me.TxtMonto1.Text)
 If Monto < 1 Or Me.txtOrigen.Text <> "" Then
    strCadena = "SELECT * FROM DocumentoCompra WHERE cPersona='" & Val(Me.txtCodPersona.Text) & "' AND seleccion='V' AND saldo>0 AND anulado='F' ORDER BY saldo DESc"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
         If Me.OptChequeSi.Value = True Then
            operacion = "CKEQUE:" + str(Me.DtcCheque.Text)
          Else
            operacion = Me.txtOperacion.Text
         End If
          monto_letras = UCase(EnLetras(Monto))
         strCadena = "SELECT * FROM Det_alm_com WHERE doc_cod='0097' AND id_usuario='" & KEY_USUARIO & "'"
         Call ConfiguraTemporal(strCadena)
         If rstTemporal.RecordCount < 1 Then
            MsgBox "Asigne una serie de Rbo Egreso para este Usuario"
            Exit Sub
         End If
         
         strCadena = "INSERT INTO movimiento_caja(doc_cod,serie,numero,tipo_trans,moneda,monto,Ingreso,Egreso,saldo,operacion,fecha_valor,fecha_sys,id_costo,glosa,cambio," & _
         "codigo_per,cPersona,descripcion_per,monto_letras,id_cuenta,Ruc) VALUES ('0097','" & rstTemporal("serie") & "','" & rstTemporal("numero") & "','E'," & _
         "'" & Me.DtcMoneda.Text & "','" & Monto & "','0','" & Monto & "','0','" & operacion & "','" & CVDate(Me.DtpValor.Value) & "','" & KEY_FECHA & "'," & _
        "'00120','" & Me.TxtGlosa.Text & "','" & Me.TxtTc.Text & "','" & Me.TxtRuc.Text & "','" & Trim(Me.txtCodPersona.Text) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(monto_letras) & "','" & Me.txtOrigen.BoundText & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        idRecibo = IdInsert("movimiento_caja")
        For i = 0 To rst.RecordCount - 1
            saldof = Monto - rst("saldo")
            If saldof > 0 Then
                depositar = rst("saldo")
                Monto = saldof
            Else
                depositar = Monto
                Monto = saldof
            End If
          If depositar > 0 Then
            monto_letras = UCase(EnLetras(depositar))
            recibo = "RBO EGRE:" + Mid(rstTemporal("serie"), 2, 4) + "-" + rstTemporal("numero")
            strCadena = "INSERT INTO mis_cuentas_det(id_cuenta,fecha,fecha_sys,tipo_trans,cPersona,Persona,glosa,monto,montoreal,tc,monto_letras,operacion,documento,IdMovimiento,recibo,id_usuario,Ruc) " & _
            " VALUES('" & Val(Me.txtOrigen.BoundText) & "','" & CVDate(Me.DtpValor.Value) & "','" & CVDate(Date) & "','E','" & Trim(Me.txtCodPersona.Text) & "'," & _
            "'" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(depositar) & "','" & Val(depositar) * -1 & "','" & Val(Me.TxtTc.Text) & "','" & monto_letras & "'," & _
            "'" & operacion & "','" & idRecibo & "','" & rst("idCompra") & "','" & recibo & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE DocumentoCompra SET saldo='" & rst("saldo") - Val(depositar) & "' WHERE idCompra='" & rst("idCompra") & "'"
            CnBd.Execute (strCadena)
             
        End If
        rst.MoveNext
    Next i
    nuevo_numero = formato_item(Val(rstTemporal("numero")) + 1, 6)
    strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & rstTemporal("serie") & "' AND doc_cod='" & rstTemporal("doc_cod") & "')"
    CnBd.Execute (strCadena)
     
    Set rstTemporal = Nothing
    End If
    'verificar si se paga con cheque
    If Me.OptChequeSi.Value = True Then
          strCadena = "SELECT * FROM Det_alm_com WHERE doc_cod='0096' AND id_usuario='" & KEY_USUARIO & "' "
          Call ConfiguraTemporal(strCadena)
          If rstTemporal.RecordCount < 1 Then
                MsgBox "Asigne una serie de Rbo Egreso para este Usuario"
            Exit Sub
          End If
            Set rst = Nothing
            strCadena = "SELECT * FROM Cheques WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_cuenta='" & Val(Me.txtOrigen.BoundText) & "'"
            Call ConfiguraRst(strCadena)
            strCadena = "UPDATE  Cheques SET monto='" & Val(Me.TxtMonto1.Text) & "',cPersona='" & Trim(Me.txtCodPersona.Text) & "',Persona='" & Trim(Me.DtcEntidad.Text) & "',fecha='" & CVDate(Me.DtpValor.Value) & "'," & _
            "estado='emitido',detalle='" & Trim(Me.TxtGlosa.Text) & "',operacion='" & Me.txtOperacion.Text & "' WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_chequera='" & Val(rst("id_chequera")) & "' AND Ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
            monto_letras = UCase(EnLetras(Val(Me.TxtMonto1.Text)))
             strCadena = "INSERT INTO OrdenPago(doc_cod,serie,numero,empresa,direccion,ruc_emp,num_cheque,cpersona,persona,entidad_financiera,fecha,cambio,glosa,monto,monto_letras)VALUES " & _
            " ('0096','" & rstTemporal("serie") & "','" & rstTemporal("numero") & "','" & Trim(KEY_EMPRESA) & "','" & Trim(KEY_DIRECCION) & "'," & _
            "'" & Trim(KEY_RUC) & "','" & Trim(Me.DtcCheque.BoundText) & "','" & Trim(Me.txtCodPersona.Text) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.txtOrigen.Text) & "','" & CVDate(Me.DtpValor.Value) & "'," & _
            "'" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtMonto1.Text) & "','" & Trim(monto_letras) & "')"
             CnBd.Execute (strCadena)
              
             nordenpago = IdInsert("OrdenPago")
             Me.TlbAcciones.Buttons(KEY_ORDENPAGO).Enabled = True
             Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
             numero_orden = formato_item(Val(rstTemporal("numero")) + 1, 6)
             strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(numero_orden) & "'  WHERE (serie='" & rstTemporal("serie") & "' AND doc_cod='0096')"
            CnBd.Execute (strCadena)
             
            Exit Sub

        Else
        
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
        Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
        Me.TlbAcciones.Buttons(KEY_ORDENPAGO).Enabled = False
        
    End If
  
 End If
  
  Exit Sub
salir:
  MsgBox "Ocurrio un Error al Grabar", vbInformation, "Mensaje para el Usuario"
  
End Sub
Private Sub Save_varios()
Dim Saldo As Single
Dim monto_letras As String
Dim Monto As Double
Dim pagado As Single
Dim glosa As String
Dim glosaITF As String
Dim rstc As New ADODB.Recordset
  On Error GoTo salir
  If (Me.txtOrigen.Text <> "" Or Val(Me.TxtMonto1.Text) <= 0) Then
    
    strCadena = "SELECT DocumentoCompra.idCompra, Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra," & _
    " DocumentoCompra.Saldo FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE seleccion='si' AND IdUsuario='" & Trim(KEY_USUARIO) & "' ORDER BY dEmisionCompra ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Monto = Val(Me.TxtMonto1.Text)
        For i = 0 To rst.RecordCount - 1
            If (Monto > 0) Then
                If (Monto > rst("saldo")) Then
                    Saldo = 0
                    pagado = rst("saldo")
                Else
                    Saldo = rst("saldo") - Monto
                    pagado = Monto
                End If
                
                strCadena = "UPDATE DocumentoCompra SET saldo='" & Val(Saldo) & "' WHERE idCompra='" & rst("idCompra") & "'"
                CnBd.Execute (strCadena)
                 
                Monto = Monto - Val(rst("saldo"))
                
            End If
            rst.MoveNext
        Next i
        
        strCadena = "SELECT     plan_contable_det.pc_codigo, mis_cuentas.id_cuenta, mis_cuentas.descripcion " & _
        "FROM mis_cuentas INNER JOIN plan_contable_det ON mis_cuentas.cuenta_ctble = plan_contable_det.pc_codigo WHERE plan_contable_det.pc_codigo='" & Trim(Me.txtOrigen.BoundText) & "'"
        rstc.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
        If rst.RecordCount < 1 Then
            MsgBox "Cuenta Caja no Asociada a Ninguna Cuenta", vbInformation, "Mensaje para el Usuario"
            Set rst = Nothing
            Exit Sub
        Else
            codigo_cuenta = rstc("id_cuenta")
            descripcion_cuenta = rstc("descripcion")
            
        End If
        
        Set rstc = Nothing
        
        strCadena = "SELECT * FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
        rstc.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
        cPersona = rstc("cPersona")
        Set rstc = Nothing
        
        strCadena = "INSERT INTO mis_cuentas_det(id_cuenta,fecha,fecha_sys,tipo_trans,cPersona,Persona,glosa,monto,montoreal,tc,documento,operacion) " & _
        " VALUES('" & Val(codigo_cuenta) & "','" & CVDate(Me.DtpValor.Value) & "','" & CVDate(Date) & "','E','" & Trim(cPersona) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtMonto1.Text) & "','" & Val(Me.TxtMonto1.Text) * -1 & "','" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Trim(Me.txtOperacion.Text) & "')"
        CnBd.Execute (strCadena)
         
        Monto = Val(Me.TxtMonto1.Text)
        monto_letras = UCase(EnLetras(Monto))
        
        If Me.OptChequeSi.Value = True Then
          strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='0003' AND doc_cod='0096') ORDER BY numero DESC"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount > 0 Then
                numero_orden = rst(0)
           Else
                numero_orden = GeneraCodigo(6)
            End If
            Set rst = Nothing
            strCadena = "SELECT * FROM Cheques WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_cuenta='" & Val(codigo_cuenta) & "'"
            Call ConfiguraRst(strCadena)
            strCadena = "UPDATE  Cheques SET monto='" & Val(Me.TxtMonto1.Text) & "',cPersona='" & Trim(cPersona) & "',Persona='" & Trim(Me.DtcEntidad.Text) & "',fecha='" & CVDate(Me.DtpValor.Value) & "'," & _
            "estado='emitido',detalle='" & Trim(comprobante) & "' WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_chequera='" & Val(rst("id_chequera")) & "'"
            CnBd.Execute (strCadena)
             
            Set rst = Nothing
            'monto_letras = UCase(EnLetras(Val(Me.TxtMonto1.Text)))
             strCadena = "INSERT INTO OrdenPago(doc_cod,serie,numero,empresa,direccion,ruc_emp,num_cheque,cpersona,persona,entidad_financiera,fecha,cambio,glosa,monto,monto_letras)VALUES " & _
            " ('0096','0003','" & Trim(numero_orden) & "','" & Trim(KEY_EMPRESA) & "','" & Trim(KEY_DIRECCION) & "'," & _
            "'" & Trim(KEY_RUC) & "','" & Trim(Me.DtcCheque.BoundText) & "','" & Trim(cPersona) & "','" & Trim(Me.DtcEntidad.Text) & "','" & Trim(Me.lblDescripcion1.Caption) & "','" & CVDate(Me.DtpValor.Value) & "'," & _
            "'" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtMonto1.Text) & "','" & Trim(monto_letras) & "')"
             CnBd.Execute (strCadena)
              
                       
            
             If MsgBox("Desea Imprimir comprobante", vbQuestion + vbYesNo, "Mensaje para el Usuario") = vbYes Then
                strCadena = "SELECT empresa, direccion, ruc_emp, serie, numero, num_cheque, entidad_financiera, cpersona, persona, fecha, cambio, glosa, monto, monto_letras " & _
                "FROM  OrdenPago WHERE doc_cod='0096' AND serie='0003' AND numero='" & Trim(numero_orden) & "'"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptOrdenPago", , App.Path + "\Reportes\")
            End If
            numero_orden = formato_item(Val(numero_orden) + 1, 6)
            strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(numero_orden) & "'  WHERE (serie='0003' AND doc_cod='0096')"
            CnBd.Execute (strCadena)
             
            Exit Sub

        Else
             strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='0003' AND doc_cod='0097') ORDER BY numero DESC"
             Call ConfiguraRst(strCadena)
             If rst.RecordCount > 0 Then
                numero_egreso = rst(0)
              Else
                numero_egreso = GeneraCodigo(6)
              End If
              Set rst = Nothing
            
            If Me.DtcMoneda.BoundText = "0001" Then
                moneda = "soles"
            Else
                moneda = "dolares"
            End If
         'monto_letras = UCase(EnLetras(Val(Me.TxtMonto1.Text)))
         
       
        
        strCadena = "INSERT INTO movimiento_caja (doc_cod,serie,numero,tipo_trans,moneda,monto,Ingreso,Egreso,saldo,operacion,fecha_valor,fecha_sys,id_costo,glosa,comprobante_rel, " & _
        "cambio,codigo_per,cPersona,descripcion_per,monto_letras,escaneo,id_cuenta,anulado,destino)VALUES ('0097','0003','" & Trim(numero_egreso) & "','E','" & Trim(moneda) & "'," & _
        "'" & Val(Me.TxtMonto1.Text) & "','0','" & Val(Me.TxtMonto1.Text) & "','0','','" & CVDate(Me.DtpValor.Value) & "','" & CVDate(Date) & "','','" & Trim(Me.TxtGlosa.Text) & "'," & _
        "'" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtRuc.Text) & "','" & Trim(cPersona) & "','" & Trim(Me.DtcEntidad.Text) & "'," & _
        "'" & Trim(monto_letras) & "','','1011','no','no')"
        CnBd.Execute (strCadena)
         
        
          
            
            
            If MsgBox("Desea Imprimir comprobante", vbQuestion + vbYesNo, "Mensaje para el Usuario") = vbYes Then
                 strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
                 "   movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor, " & _
                 "   movimiento_caja.cambio, movimiento_caja.glosa, movimiento_caja.comprobante_rel, movimiento_caja.monto, " & _
                 "   movimiento_caja.monto_letras FROM         movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
                 "   Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='0003' AND movimiento_caja.numero='" & Trim(numero_egreso) & "'"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
        End If
            nuevo_numero = formato_item(Val(numero_egreso) + 1, 6)
            strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='0003' AND doc_cod='0097')"
            CnBd.Execute (strCadena)
             
           Exit Sub
        End If
          
          
        
            
            
        End If
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
        Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
        
        Call FrmListadoFacturasCompra.facturas
    End If
  
  
  
  Exit Sub
salir:
  MsgBox "Ocurrio un Error al Grabar", vbInformation, "Mensaje para el Usuario"
End Sub

Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   If Me.txtCliente.Text = "" Then
  '      Call Resalta(Me.txtcodUuario)
   ' Else
    '    Me.DtcCCostos.SetFocus
    'End If
End If

End Sub

Private Sub txtcodUuario_KeyPress(KeyAscii As Integer)
'If (KeyAscii = 13) Then
 '   StrCadena = "SELECT * FROM Persona WHERE Per_Ruc='" & Trim(Me.txtcodUuario) & "'"
  '  Call ConfiguraRst(StrCadena)
   ' If rst.RecordCount > 0 Then
    '    Me.txtCliente.Text = rst("NombrePersona")
     '   codigo_P = rst("cPersona")
      '  Set rst = Nothing
       ' If Trim(Me.txtcodUuario.Text) = "" Then
        '    Call Resalta(Me.txtCliente)
        '    Exit Sub
        'End If
        'Me.DtcCCostos.SetFocus
    'Else
     '    Set rst = Nothing
      '  Procedencia = Nuevo
       ' Dim cod_persona As String
        'StrCadena = "SELECT * FROM Persona ORDER BY cPersona DESC"
        Call ConfiguraRst(strCadena)
        ''cod_persona = GeneraCodigo(5)
        'codigo_P = cod_persona
        'Set rst = Nothing
        'FrmDetallePersona.Show
        'FrmDetallePersona.LblCodPersona.Caption = cod_persona
        'FrmDetallePersona.TxtRuc.Text = Trim(Me.txtcodUuario.Text)
        'FrmDetallePersona.OptJuridica.Value = True
        'FrmDetallePersona.chkCliente.Value = 1
        'Call FrmDetallePersona.precionar
        'Exit Sub
    'End If
'End If
End Sub

Private Sub TxtCostos1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar3
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.TxtCostos1.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtCostos2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar4
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos2.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtCostos3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar5
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos3.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub


Private Sub txtCostos4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar6
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos4.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub





Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcIgv.SetFocus
End If
End Sub



Private Sub TxtMonto1_Change()
Dim compr As String

If Me.FrameFacturas.Visible = True Then
    If Val(Me.TxtMonto1.Text) > Val(Me.TxtTotalFacturas.Text) Then
        MsgBox "Monto Ingresado Supera El Saldo de Las Facturas", vbInformation, "Mensaje Para el Usuario"
        Call Resalta(Me.TxtMonto1)
    End If
End If

If Me.FrameComprobante.Visible = True Then
   ' If Val(Me.TxtMonto1.Text) > Val(Me.TxtTotalImporte.Text) Then
    '    MsgBox "Monto Ingresado Supera el Saldo de La Factura", vbInformation, "Mensaje Para el Usuario"
     '   Call Resalta(Me.TxtMonto1)
    'End If
End If



     strCadena = "SELECT DocumentoCompra.idCompra, Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra," & _
    " DocumentoCompra.Saldo FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE seleccion='si' AND  saldo>0 AND IdUsuario='" & Trim(KEY_USUARIO) & "' ORDER BY dEmisionCompra ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Monto = Val(Me.TxtMonto1.Text)
        glosa = "POR PAGO DE:" + Space(50)
        For i = 0 To rst.RecordCount - 1
            If (Monto > 0) Then
                
                If (Monto > rst("saldo")) Then
                    Saldo = 0
                    pagado = rst("saldo")
                Else
                    Saldo = rst("saldo") - Monto
                    pagado = Monto
                End If
                Monto = Monto - rst("saldo")
                compr = Mid(rst("doc_abrev"), 1, 3) + ":" + Right(rst("sSerie"), 3) + "-" + Right(rst("cDocumentoCompra"), 6) + "->" + str(pagado) + " - "
                glosa = glosa + compr
            End If
            rst.MoveNext
        Next i
        Me.TxtGlosa.Text = glosa
  End If

End Sub

Private Sub TxtMonto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.TxtMonto1.Text) <> 0 Then
        Me.TxtMonto1.Text = Format(Val(Me.TxtMonto1.Text), "###0.00")
        Me.txtOrigen.SetFocus
    
    End If
End If
End Sub






Private Sub TxtNaturaleza_Change()
If Trim(Me.TxtNaturaleza.Text) <> "" Then
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.TxtNaturaleza.Text) & "' AND id_plancontable='0001'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.LblDescripcion2.Caption = rst("plan_des")
    Else
        Me.LblDescripcion2.Caption = ""
    End If
    Set rst = Nothing
End If
End Sub

Private Sub TxtNaturaleza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar2
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.TxtNaturaleza.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub
Public Sub precionarCompraEgreso()
    
    strCadena = "SELECT     DocumentoCompra.idCompra, Persona.Per_Ruc, DocumentoCompra.Persona,nTotalCompra ,dEmisionCompra,moneda,saldo,nIgv  " & _
    "FROM         DocumentoCompra INNER JOIN Persona ON DocumentoCompra.cPersona = Persona.cPersona  " & _
    "WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "' "
     Call ConfiguraRst(strCadena)
        Me.FrameCCostos.Visible = True
        Me.FrmaeMonto.Visible = False
        Me.FrmCheque.Visible = False
        Me.TxtIdCompra.Text = rst("idCompra")
        Me.TxtTotalImporte.Text = Format(rst("saldo"), "###0.00")
        Me.TxtMontoFactura.Text = Format(rst("nTotalCompra"), "###0.00")
        Me.DtcMoneda.BoundText = rst("moneda")
        Me.TxtRuc.Text = rst("Per_Ruc")
        Me.DtpValor.Value = rst("dEmisionCompra")
        If rst("nIgv") > 0 Then
            Me.DtcIgv.BoundText = "si"
        Else
            Me.DtcIgv.BoundText = "no"
        End If
        strCadena = "SELECT Per_Ruc as Codigo,NombrePersona as Descripcion FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcEntidad)
        Me.TxtGlosa.Text = "CANCELACION DE COMPROBANTES"
        Call Resalta(Me.TxtNaturaleza)
End Sub

Public Sub precionar()

    strCadena = "SELECT     DocumentoCompra.idCompra, Persona.Per_Ruc, DocumentoCompra.Persona,nTotalCompra ,dEmisionCompra,moneda,saldo,nIgv  " & _
    "FROM         DocumentoCompra INNER JOIN Persona ON DocumentoCompra.cPersona = Persona.cPersona  " & _
    "WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.FrameCCostos.Visible = False
        Me.TxtIdCompra.Text = rst("idCompra")
        Me.TxtTotalImporte.Text = Format(rst("saldo"), "###0.00")
        Saldo = rst("saldo")
        Me.DtcMoneda.BoundText = rst("moneda")
        Me.TxtRuc.Text = rst("Per_Ruc")
        Me.DtpValor.Value = rst("dEmisionCompra")
        If rst("nIgv") > 0 Then
            Me.DtcIgv.BoundText = "si"
        Else
            Me.DtcIgv.BoundText = "no"
        End If
        '-
         strCadena = "SELECT Per_Ruc as Codigo,NombrePersona as Descripcion FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcEntidad)
        
        If Saldo = 0 Or Saldo < 0 Then
            MsgBox "Este Comprobante ya fue Cancelado", vbInformation, "Mensaje para el Usuario"
            Me.DtcMovimiento.SetFocus
            Exit Sub
        End If
        Call Resalta(Me.TxtMonto1)
    Else
        Me.FrameCCostos.Visible = True
        If MsgBox("Desea Ingresar como Compra ?", vbQuestion + vbYesNo, "Mensaje para el Usuario") = vbYes Then
            FrmCompras.Prender
            FrmCompras.DtcTipoDoc.BoundText = Trim(Me.TxtTD.Text)
            FrmCompras.TxtSerie.Text = Trim(Me.TxtSerie.Text)
            FrmCompras.TxtNumeroDoc.Text = Trim(Me.txtnumero)
          '  FrmCompras.TxtCodProveedor.SetFocus
            FrmCompras.Show
            Exit Sub
        Else
              Call Resalta(Me.TxtTotalImporte)
        End If
        
    End If
End Sub
Public Sub precionar_egreso()
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 6)
    strCadena = "SELECT     movimiento_caja.glosa, movimiento_caja.monto, movimiento_caja.cPersona, movimiento_caja.descripcion_per, " & _
    "Persona.Per_Ruc,movimiento_caja.codigo,movimiento_caja.moneda,movimiento_caja.fecha_valor FROM         movimiento_caja INNER JOIN Persona ON movimiento_caja.cPersona = Persona.cPersona  " & _
    "WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.txtnumero.Text) & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.FrameCCostos.Visible = True
        Me.FrmaeMonto.Visible = False
        Me.TxtIdCompra.Text = rst("codigo")
        Me.TxtTotalImporte.Text = Format(rst("monto"), "###0.00")
        Me.TxtMontoFactura.Text = Format(rst("monto"), "###0.00")
        Me.DtcMoneda.BoundText = rst("moneda")
        Me.TxtRuc.Text = rst("Per_Ruc")
        Me.DtpValor.Value = rst("fecha_valor")
        Me.FrmCheque.Visible = False
        Me.TxtGlosa.Text = rst("glosa")
        Set rst = Nothing
        strCadena = "SELECT Per_Ruc as Codigo,NombrePersona as Descripcion FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcEntidad)
        Call Resalta(Me.TxtNaturaleza)
        Set rst = Nothing
    End If
End Sub


Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
Dim Saldo As Single
If KeyAscii = 13 Then
    strCadena = "UPDATE DocumentoCompra SET seleccion='no' "
    CnBd.Execute (strCadena)
     
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 10)
    strCadena = "SELECT * FROM DocumentoCompra WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("tipo_factura") = "normal" Then
            Me.FrameIntacta.Visible = False
            Call precionar
            Call consultaCC
        Else
            Me.FrameIntacta.Visible = True
            Me.TxtMontoFactura.Text = Format(rst("nTotalCompra"), "###0.00")
            Call precionarCompraEgreso
              Call consultaCC
        End If
        Set rst = Nothing
        Exit Sub
    End If

    If Trim(Me.TxtTD.Text) = "0097" Then
        Call precionar_egreso
        Call consultaCC
        Exit Sub
    End If
    Call precionar
    Call consultaCC
End If
End Sub
Private Sub consultaCC()
Dim naturaleza As String
Dim cantidad As Integer

strCadena = "SELECT cuenta_naturaleza, cuenta, monto FROM CentroCostosDoc WHERE id_comprobante='" & Val(Me.TxtIdCompra.Text) & "' AND doc_cod='" & Trim(Me.TxtTD.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
    rst.MoveFirst
    naturaleza = rst("cuenta_naturaleza")
   
    cantidad = 1
    If cantidad <= rst.RecordCount Then
        Me.TxtCostos1.Text = rst("cuenta")
        Me.Monto1.Text = Format(rst("monto"), "###0.00")
        cantidad = cantidad + 1
    End If
    If cantidad <= rst.RecordCount Then
        rst.MoveNext
        Me.txtCostos2.Text = rst("cuenta")
        Me.Monto2.Text = Format(rst("monto"), "###0.00")
        cantidad = cantidad + 1
    End If
    If cantidad <= rst.RecordCount Then
        rst.MoveNext
        Me.txtCostos3.Text = rst("cuenta")
        Me.Monto3.Text = Format(rst("monto"), "###0.00")
        cantidad = cantidad + 1
    End If
    If cantidad <= rst.RecordCount Then
        rst.MoveNext
        Me.txtCostos4.Text = rst("cuenta")
        Me.Monto4.Text = Format(rst("monto"), "###0.00")
        cantidad = cantidad + 1
    End If
     Me.TxtNaturaleza.Text = naturaleza
     
    Set rst = Nothing
    
End If
End Sub


Private Sub TxtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtRuc)
End If
End Sub

Private Sub txtOrigen_Change()
If Trim(Me.txtOrigen.Text) <> "" Then
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.txtOrigen.BoundText) & "' AND id_plancontable='0001'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblDescripcion1.Caption = rst("plan_des")
    Else
        Me.lblDescripcion1.Caption = ""
    End If
    Set rst = Nothing
End If
End Sub

Private Sub txtOrigen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
End If
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscarPersona
End If
End Sub
Public Sub buscarPersona()
If Me.TxtRuc.Text <> "" Then
        strCadena = "SELECT Per_Ruc as Codigo,NombrePersona as Descripcion FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Call LlenaDataCombo(Me.DtcEntidad)
            Call Resalta(Me.TxtGlosa)
        Else
            Set rst = Nothing
            'Procedencia = Nuevo
            FrmDetallePersona.Show
            Exit Sub
        End If
        Set rst = Nothing
    Else
    
        Procedencia = buscar
        FrmPersona.Show
        Exit Sub
    End If
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = FormatosCeros(Me.TxtSerie.Text, 4)
    Call Resalta(Me.txtnumero)
End If
End Sub

Private Sub TxtTD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtTD.Text = "" Then
        Procedencia = buscar
        FrmComprobantes.Show
        Exit Sub
    Else
        Me.TxtTD.Text = formato_item(Me.TxtTD.Text, 4)
    
    End If
    Call Resalta(Me.TxtSerie)
End If
 
End Sub


Private Sub TxtTotalImporte_Change()
Dim rstG As New ADODB.Recordset

If Val(Me.TxtTotalImporte.Text) > 0 Then
    strCadena = "SELECT * FROM DocumentoCompra WHERE doc_cod='" & Trim(Me.TxtTD.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND cDocumentoCompra='" & Trim(Me.txtnumero.Text) & "'"
    rstG.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    If rstG.RecordCount < 1 Then
        Me.TxtMontoFactura.Text = Format(Val(Me.TxtTotalImporte.Text), "###0.00")
    End If
    Set rstG = Nothing
End If
End Sub

Private Sub TxtTotalImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtTotalImporte.Text = Format(Val(Me.TxtTotalImporte.Text), "###0.00")
    Me.DtcMoneda.SetFocus
End If
End Sub



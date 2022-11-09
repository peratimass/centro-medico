VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmComprobantesCaja 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComprobante 
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
      Left            =   1560
      TabIndex        =   58
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   15255
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   975
         Left            =   5040
         TabIndex        =   62
         Top             =   5520
         Width           =   3615
      End
      Begin VB.Frame frm_motivo_nota 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   480
         TabIndex        =   50
         Top             =   5520
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CommandButton cmdcerrarmotivonota 
            Height          =   255
            Left            =   4200
            Picture         =   "frmComprobantesCaja.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   40
            Width           =   255
         End
         Begin VB.TextBox txtmotivo_nota 
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
            Height          =   450
            Left            =   960
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            ToolTipText     =   "INGRESE UN MOTIVO"
            Top             =   400
            Width           =   3225
         End
         Begin MSDataListLib.DataCombo DtcTipoNota 
            Height          =   315
            Left            =   960
            TabIndex        =   53
            Top             =   45
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   ""
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO NOTA :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   795
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MOTIVO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   630
         End
      End
      Begin VB.Frame frm_cantidad 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   9000
         TabIndex        =   42
         Top             =   2520
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtCantidad 
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
            Left            =   1200
            TabIndex        =   48
            Top             =   1080
            Width           =   975
         End
         Begin VB.Image cmdcerrar 
            Height          =   240
            Left            =   3960
            Picture         =   "frmComprobantesCaja.frx":2EA4
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lblcantidad_original 
            BackColor       =   &H00808080&
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
            Height          =   255
            Left            =   2280
            TabIndex        =   56
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lbldescripcion 
            BackColor       =   &H00808080&
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
            Height          =   495
            Left            =   1200
            TabIndex        =   47
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblcodigo 
            BackColor       =   &H00808080&
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
            Height          =   255
            Left            =   1200
            TabIndex        =   46
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
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
            Left            =   120
            TabIndex        =   45
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION :"
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
            TabIndex        =   44
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "CODIGO :"
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
            TabIndex        =   43
            Top             =   120
            Width           =   630
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdquitarItem 
         Height          =   765
         Left            =   13560
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1349
         BTYPE           =   5
         TX              =   "QUITAR ITEM"
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
         MICON           =   "frmComprobantesCaja.frx":5D48
         PICN            =   "frmComprobantesCaja.frx":5D64
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrardetalle 
         Height          =   285
         Left            =   14640
         TabIndex        =   27
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   5
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmComprobantesCaja.frx":81AE
         PICN            =   "frmComprobantesCaja.frx":81CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGenerarNotaCredito 
         Height          =   645
         Left            =   8760
         TabIndex        =   28
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1138
         BTYPE           =   5
         TX              =   "GENERAR N.CREDITO"
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
         MICON           =   "frmComprobantesCaja.frx":B07E
         PICN            =   "frmComprobantesCaja.frx":B09A
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
         Height          =   645
         Left            =   10980
         TabIndex        =   29
         Top             =   5760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1138
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
         MICON           =   "frmComprobantesCaja.frx":EDDB
         PICN            =   "frmComprobantesCaja.frx":EDF7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdModificarComprobante 
         Height          =   765
         Left            =   13560
         TabIndex        =   40
         Top             =   3840
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1349
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmComprobantesCaja.frx":113C8
         PICN            =   "frmComprobantesCaja.frx":113E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
         Height          =   4095
         Left            =   480
         TabIndex        =   41
         Top             =   1340
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   7223
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
      Begin VitekeySoft.ChameleonBtn cmdNotaCredito 
         Height          =   765
         Left            =   13560
         TabIndex        =   49
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1349
         BTYPE           =   5
         TX              =   "GEN. N.CREDITO"
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
         MICON           =   "frmComprobantesCaja.frx":13A1D
         PICN            =   "frmComprobantesCaja.frx":13A39
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblVendedor 
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
         Left            =   10200
         TabIndex        =   39
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "VENDEDOR :"
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
         Left            =   8880
         TabIndex        =   38
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblcomprobante 
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
         Left            =   1860
         TabIndex        =   37
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPROBANTE :"
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
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lblfecha 
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
         Left            =   10140
         TabIndex        =   35
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblrazonsocial 
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
         Left            =   1860
         TabIndex        =   34
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label lblruc 
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
         Left            =   1860
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA EMISION :"
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
         Left            =   8580
         TabIndex        =   32
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RAZON SOCIAL :"
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
         Left            =   480
         TabIndex        =   31
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DNI / RUC :"
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
         Left            =   540
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.TextBox txtid_manifiesto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   15600
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtManifiesto 
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
      Height          =   300
      Left            =   13200
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox chk_credito 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CREDITO"
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
      Left            =   11640
      TabIndex        =   17
      Top             =   840
      Width           =   1335
   End
   Begin VB.CheckBox chk_contado 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "CONTADO"
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
      Left            =   10200
      TabIndex        =   16
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtcliente 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.CheckBox chk_vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "VENDEDOR :"
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
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
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
   Begin VitekeySoft.ChameleonBtn cmdconsultar 
      Height          =   675
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1191
      BTYPE           =   5
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
      MICON           =   "frmComprobantesCaja.frx":161C0
      PICN            =   "frmComprobantesCaja.frx":161DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   115081217
      CurrentDate     =   42236
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   285
      Left            =   16680
      TabIndex        =   5
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      BTYPE           =   5
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmComprobantesCaja.frx":187C1
      PICN            =   "frmComprobantesCaja.frx":187DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdGenerarasiento 
      Height          =   885
      Left            =   15600
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1561
      BTYPE           =   5
      TX              =   "GENERAR ASIENTO CAJA BANCOS"
      ENAB            =   0   'False
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
      MICON           =   "frmComprobantesCaja.frx":1B691
      PICN            =   "frmComprobantesCaja.frx":1B6AD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCobroGlobal 
      Height          =   885
      Left            =   15600
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1561
      BTYPE           =   5
      TX              =   "G.COBRO GLOBAL"
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
      MICON           =   "frmComprobantesCaja.frx":1DE34
      PICN            =   "frmComprobantesCaja.frx":1DE50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdQuitarLista 
      Height          =   885
      Left            =   15600
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1561
      BTYPE           =   5
      TX              =   "QUITAR DE LISTA"
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
      MICON           =   "frmComprobantesCaja.frx":21B91
      PICN            =   "frmComprobantesCaja.frx":21BAD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmComprobantesCaja.frx":21EC7
      Height          =   6615
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   11668
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12582912
      GridColorFixed  =   8388608
      GridColorUnpopulated=   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   2
      AllowUserResizing=   1
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   17520
      TabIndex        =   12
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   115081217
      CurrentDate     =   42236
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   315
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   115081217
      CurrentDate     =   42236
   End
   Begin MSDataListLib.DataCombo DtcManifiesto 
      Height          =   330
      Left            =   10200
      TabIndex        =   14
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
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
   Begin VitekeySoft.ChameleonBtn CmdConsultarManifiesto 
      Height          =   345
      Left            =   14040
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   5
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
      MICON           =   "frmComprobantesCaja.frx":21EDC
      PICN            =   "frmComprobantesCaja.frx":21EF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpCaja 
      Height          =   315
      Left            =   7320
      TabIndex        =   20
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   115081217
      CurrentDate     =   42236
   End
   Begin VitekeySoft.ChameleonBtn cmdCobranza 
      Height          =   950
      Left            =   15600
      TabIndex        =   21
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1667
      BTYPE           =   5
      TX              =   "REPORT. PENDIENTES DE COBRO"
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
      MICON           =   "frmComprobantesCaja.frx":244DD
      PICN            =   "frmComprobantesCaja.frx":244F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRechazos 
      Height          =   885
      Left            =   15600
      TabIndex        =   59
      Top             =   6840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1561
      BTYPE           =   5
      TX              =   "REPORTE  RECHAZOS"
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
      MICON           =   "frmComprobantesCaja.frx":26ACA
      PICN            =   "frmComprobantesCaja.frx":26AE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar prog_indicador 
      Height          =   150
      Left            =   15600
      TabIndex        =   60
      Top             =   3435
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VitekeySoft.ChameleonBtn cmdReporteCobroRealizado 
      Height          =   945
      Left            =   15600
      TabIndex        =   61
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1667
      BTYPE           =   5
      TX              =   "REPORT. COBROS REALIZADOS"
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
      MICON           =   "frmComprobantesCaja.frx":26E00
      PICN            =   "frmComprobantesCaja.frx":26E1C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMPROBANTE :"
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
      Left            =   240
      TabIndex        =   57
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MODULO ALMACEN CAJA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   12120
      TabIndex        =   24
      Top             =   120
      Width           =   2580
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MANIFIESTO :"
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
      TabIndex        =   22
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "FECHA CAJA:"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   600
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   975
      Left            =   8880
      Top             =   720
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA     :"
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
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLIENTE  :"
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
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   8505
      Left            =   0
      Top             =   0
      Width           =   17265
   End
End
Attribute VB_Name = "frmComprobantesCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents oEditFlex As clsFlex
Attribute oEditFlex.VB_VarHelpID = -1

Private Sub cmdCerrar_Click()
Me.frm_cantidad.Visible = False
End Sub

Private Sub cmdcerrardetalle_Click()
Me.frmdetalle.Visible = False

Me.frm_motivo_nota.Visible = False
Me.cmdGenerarNotaCredito.Visible = False


End Sub

Private Sub cmdCobranza_Click()

strCadena = "call ADM_reporte_standart('2','" & Me.DtcManifiesto.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptReporteCobranza", , App.Path + "\Reportes\")


End Sub

Private Sub cmdCobroGlobal_Click()
If MsgBox("Esta Seguro de Generar los COBROS", vbQuestion + vbYesNo) = vbYes Then
    
    Me.prog_indicador.Min = 0
    Me.prog_indicador.Max = Me.MSHFlexGrid1.Rows
    For i = 0 To Me.MSHFlexGrid1.Rows - 1
        If Val(Me.MSHFlexGrid1.TextMatrix(i, 6)) > 0 Then
            Call generar_caja(Me.MSHFlexGrid1.TextMatrix(i, 0), Me.MSHFlexGrid1.TextMatrix(i, 6))
        End If
        
        DoEvents
        Me.prog_indicador.Value = i
    Next i
    
    
    MsgBox "Se han Generado los asientos de Caja" + Chr(13) + "Correctamente !!", vbInformation
    
    
End If
End Sub

Private Sub cmdConsultar_Click()


If Me.chk_vendedor.Value = 0 Then
    strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE fecha_emision>='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE fecha_emision>='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_vendedor='" & Me.DtcVendedor.BoundText & "' AND ruc='" & KEY_RUC & "'"
End If

Call actualizar(Me.MSHFlexGrid1)

Me.txtid_manifiesto.Text = 0


End Sub
Private Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_obsequio As Single
strCadena = "SELECT * FROM view_detalle_venta WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    
    Grilla.Rows = 0
    in_obsequio = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 0
           'Grilla.ColAlignment(4) = 7
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_obsequio = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
               in_producto = ""
               in_unidad = ""
               If rst("cantidad") = 0 Then
                  in_cantidad = ""
                Else
                  in_cantidad = Format(rst("cantidad"), "###0.00")
               End If
               
               If rst("precio") = 0 Then
                  in_precio = ""
                Else
                  in_precio = Format(rst("precio"), "###0.00")
               End If
               
            Else
              in_producto = rst("id_producto")
              in_unidad = rst("abreviatura")
              in_cantidad = Format(rst("cantidad"), "###0.00")
              in_precio = Format(rst("precio"), "###0.00")
            End If
            
            
            
            Fila = rst("id_detalle_venta") & vbTab & in_producto & vbTab & rst("detalle") & vbTab & in_unidad & vbTab & in_cantidad & vbTab & in_precio & vbTab & Format(rst("total"), "###0.00")
            Grilla.AddItem Fila
            If (Trim(rst("id_igv")) = "no") Then
                            texonerado = texonerado + rst("total")
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0FFFF
                            Next k
             Else
                            tafecto = tafecto + rst("total")
             End If
             If rst("obsequio") = "si" Then
                in_obsequio = in_obsequio + in_precio * in_cantidad
                For k = 3 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
             
            
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Me.frmdetalle.Visible = True

Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Sub llenar(ByVal id_venta As Double)

strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Me.lblComprobante.Caption = rst("documento")
Me.lblruc.Caption = rst("id_cliente")

Me.lblrazonsocial.Caption = rst("ncliente")
Me.lblfecha.Caption = str(rst("fecha_emision"))
Me.lblVendedor.Caption = get_persona(rst("id_vendedor"))

End Sub

Private Sub CmdConsultarManifiesto_Click()
  
 
 in_criterio = ""
 If Me.chk_contado.Value = 1 Then
    in_criterio = " and id_forma_pago='01'"
 End If
 
 If Me.chk_credito.Value = 1 Then
    in_criterio = in_criterio & " and id_forma_pago='02'"
 End If
  
 
 strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE id_manifiesto='" & Me.DtcManifiesto.BoundText & "' AND ruc='" & KEY_RUC & "'" & in_criterio
 Call actualizar(Me.MSHFlexGrid1)
 Me.txtid_manifiesto.Text = Val(Me.DtcManifiesto.BoundText)
 
 
 
 
 
 
End Sub

Private Sub cmdGenerarasiento_Click()

Call generar_caja(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0), Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)))

If Me.chk_vendedor.Value = 0 Then
    strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE fecha_emision='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE fecha_emision='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and id_vendedor='" & Me.DtcVendedor.BoundText & "' AND ruc='" & KEY_RUC & "'"
End If

Call actualizar(Me.MSHFlexGrid1)

End Sub
Private Sub generar_caja(ByVal in_venta As String, ByVal in_monto_caja As Double)
Dim in_mis_cuentas_det As String

strCadena = "SELECT* FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstA(strCadena)
    


strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       in_documento = rstA("documento")
       in_glosa = "PAGO :" & in_documento
       in_flujo = "1CIX000000000078"
       in_documento = rstA("documento")
       in_doc = rstA("id_doc")
       in_cliente = rstA("id_cliente")
       'Solo van a pasar los que tienen Saldo
       
       in_mis_cuentas_det = procesar_transaccion_venta(rstK("id_forma_pago"), KEY_ALM, get_cuenta_pago(rstK("id_forma_pago")), Format(Me.DtpCaja.Value, "YYYY-mm-dd"), "00001", Trim(in_cliente), get_persona(in_cliente), in_glosa, in_monto_caja, "0", in_venta, "0", in_documento, Val(KEY_CAMBIO), rstK("id_tarjeta_operacion"), "1CIX000000000174", in_flujo, KEY_USUARIO, in_doc, KEY_RUC)
       Call put_realizar_pago(in_venta, in_venta, Abs(in_monto_caja), in_doc, Val(KEY_CAMBIO), Val(in_mis_cuentas_det))
       
       
       strCadena = "call sp_insertar_transaccion_caja('" & Val(in_mis_cuentas_det) & "')"
       CnBd.Execute (strCadena)
       
       strCadena = "CALL ADM_ventas_utilidades('1','" & in_venta & "','" & Me.DtcManifiesto.BoundText & "','" & in_monto_caja & "','','','','','','" & KEY_RUC & "')"
       Call ConfiguraRstAux(strCadena)
       
       rstK.MoveNext
   Next i
   
   
   
    strCadena = "SELECT(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstA(strCadena)
    If rstA("saldo") = 0 Then
        strCadena = "UPDATE movimiento_venta SET asiento_caja='si' WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
    End If

   
   
   
   Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 8) = "GENERADO"
   
  
   
End If

End Sub

Private Sub load_item(ByVal in_detalle As String)
Me.frm_cantidad.Visible = True
Me.lblcodigo.Caption = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)
Me.lbldescripcion.Caption = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 2)
Me.txtCantidad.Text = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 4)
Me.lblcantidad_original.Caption = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 4)
Me.lblcantidad_original.Tag = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 5)

End Sub

Private Sub cmdModificarComprobante_Click()
Call load_item(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0))
End Sub

Private Sub cmdNotaCredito_Click()
If MsgBox("Desea realizar Nota Credito", vbQuestion + vbYesNo) = vbYes Then
   Me.frm_motivo_nota.Visible = True
   Me.cmdGenerarNotaCredito.Visible = True
   Call load_tipo_nota
End If
End Sub


Private Sub load_tipo_nota()
strCadena = "SELECT id_tipo_nota as Codigo,descripcion as Descripcion FROM tipo_nota_credito ORDER BY id_tipo_nota"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoNota)
Me.frm_motivo_nota.Visible = True
End Sub

Private Sub cmdquitarItem_Click()
Me.HfDetalle.RemoveItem (Me.HfDetalle.Row)
End Sub

Private Sub cmdQuitarLista_Click()
If Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)) > 0 Then
    
    'strCadena = "call put_quitar_comprobante_pago('" & Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)) & "','" & Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 5)) & "','" & Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6)) & "','" & KEY_RUC & "')"
    'CnBd.Execute (strCadena)
    Me.MSHFlexGrid1.RemoveItem (Me.MSHFlexGrid1.Row)
    Call get_total
    Call get_total_pagar
End If
End Sub

Private Sub cmdReporteCobroRealizado_Click()
strCadena = "call ADM_reporte_standart('3','" & Me.DtcManifiesto.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptReporteCobranza", , App.Path + "\Reportes\")

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

Me.DTPicker1.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

Me.DtpCaja.Value = KEY_FECHA


'strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE fecha_emision='" & KEY_FECHA & "' AND ruc='" & KEY_RUC & "'"
'Call actualizar(Me.MSHFlexGrid1)



strCadena = "SELECT id_manifiesto as Codigo,manifiesto as Descripcion FROM view_manifiesto_numero WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcManifiesto)

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE  id_personal='si' and habilitado='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)






Set oEditFlex = New clsFlex
With oEditFlex
    ' edita al hacer un clic
    .EditMode = eClick
      
    ' inicia
    Call oEditFlex.Iniciar(Me.MSHFlexGrid1, Me, Me.DTPicker2)
    
    'Configura las columnas
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    ' columna 0 ( no se puede editar)
    Call .SetColumnas(0, enumber, True)
    ' columna 1 , tipo string,  se puede editar
    Call .SetColumnas(1, enumber, True)
    ' columna 2 , tipo numerico,  se puede editar
    Call .SetColumnas(2, estring, True)
    ' columna 3 , tipo moneda,  se puede editar
    Call .SetColumnas(3, estring, True)
    
    ' columna 4 , tipo boolean,  se puede editar
    Call .SetColumnas(4, estring, True)
    
    ' columna 5 , tipo fecha,  se puede editar
    Call .SetColumnas(5, estring, True)
    Call .SetColumnas(6, estring, False)
     Call .SetColumnas(7, estring, True)
     Call .SetColumnas(8, estring, True)
  End With
  


  
  

End Sub



Public Sub actualizar(ByVal Grilla As MSHFlexGrid)
Dim Anulado As String

Dim in_total As Double
Dim in_monto_pagar As Double


Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then

    Grilla.Rows = 0
    
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 1400
       Next
      
        
        cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "SALDO" & vbTab & "MONTO CAJA" & vbTab & "FORMA PAGO" & vbTab & "ASIENTO CAJA"
        Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstL.MoveFirst
        in_total = 0
        in_monto_pagar = 0
        
        For i = 0 To rstL.RecordCount - 1
            If rstL("asiento_caja") = "si" Then
                asiento_caja = "GENERADO"
            Else
                asiento_caja = "SIN GENERAR"
            End If
            
            Fila = rstL("id_venta") & vbTab & Format(rstL("fecha_emision"), "dd-mm-YYYY") & vbTab & rstL("documento") & vbTab & rstL("id_cliente") & vbTab & rstL("ncliente") & vbTab & Format(rstL("total"), "#,##0.00") & vbTab & Format(rstL("monto_pago"), "#,##0.00") & vbTab & rstL("forma_pago") & vbTab & asiento_caja
            Grilla.AddItem Fila
            
            in_total = in_total + rstL("total")
            in_monto_pagar = in_monto_pagar + rstL("monto_pago")
            
            
            
            If rstL("asiento_caja") = "no" Then
                    Grilla.col = 8
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
            End If
                
                
                rstL.MoveNext
                
                
                
        Next i
        
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_total, "#,##0.00") & vbTab & Format(in_monto_pagar, "#,##0.00")
        Grilla.AddItem Fila
        
        For j = 5 To 6
                    Grilla.col = j
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H80FF&
        Next j
            

        ' nueva instancia del módulo
            
    
        
  
End Sub
Private Sub get_total()
Dim in_acumulado As Double
in_acumulado = 0
For i = 0 To Me.MSHFlexGrid1.Rows - 2
    If Val(Me.MSHFlexGrid1.TextMatrix(i, 0)) > 0 Then
        in_acumulado = in_acumulado + Format(Me.MSHFlexGrid1.TextMatrix(i, 5), "###0.00")
    End If
Next i

Me.MSHFlexGrid1.TextMatrix(i, 5) = Format(in_acumulado, "#,##0.00")


End Sub

Private Sub get_total_pagar()
Dim in_acumulado As Double
in_acumulado = 0
For i = 0 To Me.MSHFlexGrid1.Rows - 2
    If Val(Me.MSHFlexGrid1.TextMatrix(i, 0)) > 0 Then
    in_acumulado = in_acumulado + Format(Me.MSHFlexGrid1.TextMatrix(i, 6), "###0.00")
    End If
Next i

Me.MSHFlexGrid1.TextMatrix(i, 6) = Format(in_acumulado, "#,##0.00")


End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 And Me.frm_motivo_nota.Visible = True Then
   
   Me.cmdModificarComprobante.Visible = True
   Me.cmdquitarItem.Visible = True
Else
    Me.cmdModificarComprobante.Visible = False
   Me.cmdquitarItem.Visible = False
End If
End Sub

Private Sub MSHFlexGrid1_DblClick()



Call llenar(Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)))
Call llenarGrid_Comprobante(Me.HfDetalle, Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)))
End Sub

' evento para validar el valor de la celda cuando cambia
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub oEditFlex_Validar(IRowActual As Long, _
                              IColActual As Integer, _
                              CellValue As Variant, _
                                  Cancelar As Boolean)
    
    '''''''''''''''''''''''''''''''''''''''''''
    'CellValue :  Valor de la celda editada
    'IRowActual : indice de la fila
    'IColActual : Indice de la columna
    'Cancelar : Si se coloca True, no se modifica el valor y se reestablece
    '''''''''''''''''''''''''''''''''''''''''''
    
    ' índice de la columna
    Select Case IColActual
        
        ' Id
        Case 0
            MsgBox "Esta celda no se puede editar", vbInformation
        ' Columna string
        Case 1
            If Len(CStr(CellValue)) < 4 Then
               MsgBox "El texto debe ser mayor a 4 caracteres", vbExclamation
                Cancelar = True
            End If
        
        ' columna Solonúmero
        Case 2
            If CLng(CellValue) > 10 Then
               MsgBox "Debe ingresar un número menor a 10", vbExclamation
               Cancelar = True
            End If
        'Columna moneda
        Case 3
            If CCur(CellValue) > 2000 Then
               MsgBox "El monto debe ser menor a $ 2000", vbExclamation
               Cancelar = True
            End If
        
        'Columna Boolean
        Case 4
            If IRowActual < 4 Then
               If CellValue = True Then
                  MsgBox "El valor de las primeras tres filas debe ser No. No se cambiará el dato", vbInformation
                  Cancelar = True
               End If
            End If
        
        ' Columna fecha
        Case 5
            If CDate(CellValue) < Date Then
               MsgBox "La fecha debe ser mayor a la fecha actual", vbExclamation
               Cancelar = True
            End If
            
         Case 6
            If Val(CellValue) < 0 Then
               MsgBox "Monto tiene que ser mayor que 0", vbExclamation
               Cancelar = True
            End If
    End Select
    
    
    ' si no se cancela ...
    If Cancelar = False Then
        
    
    ' nuevo recordset
        
        Dim i As Integer
        Dim id As Long
        
        ' columna con el ID
        id = Val(MSHFlexGrid1.TextMatrix(IRowActual, 0))
        If Val(id) > 0 Then
           strCadena = "call put_monto_pago('" & Val(id) & "','" & Val(CellValue) & "')"
           CnBd.Execute (strCadena)
           'Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 6) = Format(Val(CellValue), "#,##0.00")
        
       End If
       Call get_total_pagar
    End If


End Sub



Private Sub MSHFlexGrid1_SelChange()
If Val(Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, 0)) > 0 Then
   Me.cmdGenerarasiento.Enabled = True
Else
   Me.cmdGenerarasiento.Enabled = False
End If
End Sub




Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txtCantidad.Text) <= Val(Me.lblcantidad_original.Caption) Then
        Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 4) = Me.txtCantidad.Text
        Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 6) = Val(Me.txtCantidad.Text) * Val(Me.lblcantidad_original.Tag)
    Else
        MsgBox "Cantidad No tiene que Superar al Original", vbInformation
    End If
    Me.frm_cantidad.Visible = False
End If
End Sub

Private Sub txtComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_listado_comprobantes_caja WHERE documento LIKE '%" & Trim(Me.txtComprobante.Text) & "%'  AND ruc='" & KEY_RUC & "'"
    Call actualizar(Me.MSHFlexGrid1)
    Me.txtid_manifiesto.Text = 0
End If
End Sub

Private Sub txtManifiesto_Change()
strCadena = "SELECT id_manifiesto as Codigo,manifiesto as Descripcion FROM view_manifiesto_numero WHERE manifiesto LIKE '%" & Trim(Me.txtManifiesto.Text) & "%' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcManifiesto)
End Sub

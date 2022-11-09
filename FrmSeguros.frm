VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmSeguros 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   16815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmempleadoras 
      BackColor       =   &H00FFFFFF&
      Height          =   8100
      Left            =   6360
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame frmdescuento 
         BackColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   960
         TabIndex        =   70
         Top             =   2160
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txt_id_seguro_empleadora 
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
            Left            =   4320
            TabIndex        =   77
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtrazon_emp 
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
            Left            =   1680
            TabIndex        =   75
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox txtruc_emp 
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
            Left            =   1680
            TabIndex        =   74
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtdescuento_emp 
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
            Left            =   1680
            TabIndex        =   72
            Top             =   1200
            Width           =   1815
         End
         Begin VitekeySoft.ChameleonBtn cmdcerrar_factor 
            Height          =   480
            Left            =   3360
            TabIndex        =   93
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   847
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
            MICON           =   "FrmSeguros.frx":0000
            PICN            =   "FrmSeguros.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar_factor 
            Height          =   480
            Left            =   1680
            TabIndex        =   84
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   847
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmSeguros.frx":3043
            PICN            =   "FrmSeguros.frx":305F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RAZON SOCIAL :"
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
            Height          =   210
            Left            =   315
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC :"
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
            Height          =   210
            Left            =   1125
            TabIndex        =   73
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCUENTO (% )  :"
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
            Height          =   210
            Left            =   120
            TabIndex        =   71
            Top             =   1320
            Width           =   1410
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   1920
         TabIndex        =   69
         Top             =   7560
         Width           =   2775
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfEmpleadoras 
         Height          =   6735
         Left            =   360
         TabIndex        =   66
         Top             =   720
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11880
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
      Begin VitekeySoft.ChameleonBtn cmd_salir_emp 
         Height          =   855
         Left            =   8040
         TabIndex        =   87
         Top             =   3450
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
         MICON           =   "FrmSeguros.frx":66A7
         PICN            =   "FrmSeguros.frx":66C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_eliminar_emp 
         Height          =   855
         Left            =   8040
         TabIndex        =   88
         Top             =   2565
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
         MICON           =   "FrmSeguros.frx":6AB3
         PICN            =   "FrmSeguros.frx":6ACF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdagregar_emp 
         Height          =   855
         Left            =   8040
         TabIndex        =   89
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "AGREGAR"
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
         MICON           =   "FrmSeguros.frx":8F19
         PICN            =   "FrmSeguros.frx":8F35
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_editar_emp 
         Height          =   855
         Left            =   8040
         TabIndex        =   90
         Top             =   1635
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
         MICON           =   "FrmSeguros.frx":9387
         PICN            =   "FrmSeguros.frx":93A3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR EMPRESA:"
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
         Height          =   210
         Left            =   360
         TabIndex        =   68
         Top             =   7560
         Width           =   1425
      End
      Begin VB.Label lblempresa_emp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC :"
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
         Left            =   360
         TabIndex        =   67
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.Frame frmdetalle 
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
      Height          =   8115
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   15375
      Begin VitekeySoft.ChameleonBtn cmdvisualizar 
         Height          =   350
         Left            =   5160
         TabIndex        =   94
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BTYPE           =   3
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSeguros.frx":C679
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chk_pago_servicio 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AFECTA A SERVICIOS"
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
         Height          =   280
         Left            =   9840
         TabIndex        =   80
         Top             =   6600
         Width           =   2535
      End
      Begin VB.CheckBox chkconfiguracion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CONFIGURACION ESPECIAL"
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
         Height          =   280
         Left            =   9840
         TabIndex        =   78
         Top             =   6240
         Width           =   2535
      End
      Begin VB.CheckBox chkvincular 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "VINCULAR :"
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
         Left            =   8550
         TabIndex        =   56
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtvincular 
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
         Left            =   9840
         TabIndex        =   54
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtcodaseguradora 
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
         Left            =   9840
         TabIndex        =   51
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CheckBox chkhabilitado 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "HABILITADO."
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
         Height          =   280
         Left            =   9840
         TabIndex        =   49
         Top             =   5880
         Width           =   2535
      End
      Begin VB.TextBox txtdominio 
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
         Left            =   2340
         TabIndex        =   47
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox txtmail 
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
         Left            =   2340
         TabIndex        =   45
         Top             =   1440
         Width           =   5775
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "UBIGUEO"
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
         Height          =   1695
         Left            =   2340
         TabIndex        =   40
         Top             =   6360
         Width           =   5775
         Begin VB.TextBox TxtDistrito 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            Height          =   330
            Left            =   4560
            MaxLength       =   80
            TabIndex        =   41
            Top             =   1200
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DtcDistrito 
            Height          =   330
            Left            =   1650
            TabIndex        =   9
            Top             =   1185
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSDataListLib.DataCombo DtcDepartamento 
            Height          =   330
            Left            =   1650
            TabIndex        =   7
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
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
         Begin MSDataListLib.DataCombo DtcProvincia 
            Height          =   330
            Left            =   1650
            TabIndex        =   8
            Top             =   765
            Width           =   2895
            _ExtentX        =   5106
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
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PROVINCIA :"
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
            Height          =   210
            Left            =   555
            TabIndex        =   44
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lbldepartamento 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEPARTAMENTO :"
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
            Height          =   210
            Left            =   120
            TabIndex        =   43
            Top             =   405
            Width           =   1395
         End
         Begin VB.Label lbldistrito 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DISTRITO :"
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
            Height          =   210
            Left            =   675
            TabIndex        =   42
            Top             =   1245
            Width           =   855
         End
      End
      Begin VB.TextBox txtid_seguro 
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
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TELEFONOS REFERENCIA."
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
         Height          =   2775
         Left            =   8400
         TabIndex        =   29
         Top             =   360
         Width           =   6615
         Begin VB.Frame frmtelefonos 
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
            Height          =   2055
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Visible         =   0   'False
            Width           =   6375
            Begin VB.TextBox txtreferencia 
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
               Left            =   1455
               MaxLength       =   10
               TabIndex        =   11
               Top             =   1560
               Width           =   1815
            End
            Begin VB.TextBox TxtFono 
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
               Left            =   1455
               MaxLength       =   10
               TabIndex        =   10
               Top             =   1140
               Width           =   1815
            End
            Begin MSDataListLib.DataCombo DtcArea 
               Height          =   330
               Left            =   135
               TabIndex        =   34
               Top             =   600
               Width           =   3135
               _ExtentX        =   5530
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
            Begin VB.PictureBox cmdagregartelefono 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   3375
               ScaleHeight     =   330
               ScaleWidth      =   1275
               TabIndex        =   35
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "REFERENCIA :"
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
               Height          =   210
               Left            =   120
               TabIndex        =   38
               Top             =   1620
               Width           =   1065
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NUMERO :"
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
               Height          =   210
               Left            =   120
               TabIndex        =   37
               Top             =   1200
               Width           =   825
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TELEFONO :"
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
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   915
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTelefono 
            Height          =   2055
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   3625
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
         Begin VB.PictureBox cmdagregar 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            ScaleHeight     =   285
            ScaleWidth      =   795
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
         Begin VB.PictureBox cmdelete 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            ScaleHeight     =   285
            ScaleWidth      =   795
            TabIndex        =   32
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
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
         ForeColor       =   &H00800000&
         Height          =   3735
         Left            =   2340
         TabIndex        =   23
         Top             =   2280
         Width           =   5775
         Begin VB.TextBox txtucin 
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
            Left            =   2640
            TabIndex        =   63
            Top             =   2660
            Width           =   2055
         End
         Begin VB.TextBox txtfarmacia 
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
            Left            =   2640
            TabIndex        =   61
            Top             =   3080
            Width           =   2055
         End
         Begin VB.TextBox txtuci 
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
            Left            =   2640
            TabIndex        =   59
            Top             =   2260
            Width           =   2055
         End
         Begin VB.TextBox txtcama 
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
            Left            =   2640
            TabIndex        =   57
            Top             =   1840
            Width           =   2055
         End
         Begin VB.TextBox txteps 
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
            Left            =   2640
            TabIndex        =   6
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox txtvalorconsulta 
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
            Left            =   2640
            TabIndex        =   5
            Top             =   1040
            Width           =   2055
         End
         Begin VB.TextBox txtfactoraseguradora 
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
            Left            =   2640
            TabIndex        =   4
            Top             =   640
            Width           =   2055
         End
         Begin VB.TextBox TxtDescuentoproducto 
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
            Left            =   2640
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UCIN :"
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
            Height          =   210
            Left            =   1965
            TabIndex        =   64
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FARMACIA :"
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
            Height          =   210
            Left            =   1545
            TabIndex        =   62
            Top             =   3120
            Width           =   900
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UCI :"
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
            Height          =   210
            Left            =   2085
            TabIndex        =   60
            Top             =   2160
            Width           =   360
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CAMA :"
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
            Height          =   210
            Left            =   1905
            TabIndex        =   58
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EPS :"
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
            Height          =   210
            Left            =   2085
            TabIndex        =   27
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VALOR CONSULTA :"
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
            Height          =   210
            Left            =   990
            TabIndex        =   26
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FACTOR ASEGURADORA :"
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
            Height          =   210
            Left            =   525
            TabIndex        =   25
            Top             =   600
            Width           =   1920
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCUENTOS PRODUCTOS:"
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
            Height          =   210
            Left            =   390
            TabIndex        =   24
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox txtdireccion 
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
         Left            =   2340
         TabIndex        =   2
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtruc 
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
         Left            =   2340
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtdetalle 
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
         Height          =   645
         Left            =   9840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   5040
         Width           =   5175
      End
      Begin VB.TextBox TxtDescripcion 
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
         Left            =   2340
         TabIndex        =   1
         Top             =   700
         Width           =   5775
      End
      Begin MSDataListLib.DataCombo dtcVinculado 
         Height          =   330
         Left            =   9840
         TabIndex        =   53
         Top             =   3480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   600
         Left            =   12480
         TabIndex        =   91
         Top             =   7320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
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
         MICON           =   "FrmSeguros.frx":C695
         PICN            =   "FrmSeguros.frx":C6B1
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
         Height          =   600
         Left            =   10560
         TabIndex        =   92
         Top             =   7320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSeguros.frx":F6D8
         PICN            =   "FrmSeguros.frx":F6F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURACION:"
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
         Height          =   210
         Left            =   8400
         TabIndex        =   79
         Top             =   6240
         Width           =   1365
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR VINC:"
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
         Height          =   210
         Left            =   8685
         TabIndex        =   55
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD. ASEGURAD:"
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
         Height          =   210
         Left            =   8445
         TabIndex        =   52
         Top             =   4680
         Width           =   1320
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO :"
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
         Height          =   210
         Left            =   9060
         TabIndex        =   50
         Top             =   5880
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOMINIO :"
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
         Height          =   210
         Left            =   1200
         TabIndex        =   48
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-MAIL :"
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
         Height          =   210
         Left            =   1410
         TabIndex        =   46
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARAMETROS :"
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
         Height          =   210
         Left            =   900
         TabIndex        =   28
         Top             =   3840
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION FISCAL :"
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
         Height          =   210
         Left            =   540
         TabIndex        =   22
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC :"
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
         Height          =   210
         Left            =   1635
         TabIndex        =   21
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblSeguro 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   20
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Height          =   210
         Left            =   8565
         TabIndex        =   19
         Top             =   5280
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION SEGURO :"
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
         Height          =   210
         Left            =   225
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.TextBox TXTBUSCAR 
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
      Left            =   1920
      TabIndex        =   13
      Top             =   7710
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2670
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
            Picture         =   "FrmSeguros.frx":12D3C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":13190
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":134B0
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":13904
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":13D58
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":14078
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":14398
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":146B8
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSeguros.frx":149D8
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSeguros 
      Height          =   7095
      Left            =   240
      TabIndex        =   14
      Top             =   390
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   12515
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   15720
      TabIndex        =   81
      Top             =   3855
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
      MICON           =   "FrmSeguros.frx":14CF8
      PICN            =   "FrmSeguros.frx":14D14
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   855
      Left            =   15720
      TabIndex        =   82
      Top             =   2970
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
      MICON           =   "FrmSeguros.frx":15104
      PICN            =   "FrmSeguros.frx":15120
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEditar 
      Height          =   855
      Left            =   15720
      TabIndex        =   83
      Top             =   1125
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
      MICON           =   "FrmSeguros.frx":1756A
      PICN            =   "FrmSeguros.frx":17586
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
      Left            =   15720
      TabIndex        =   85
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
      MICON           =   "FrmSeguros.frx":178A0
      PICN            =   "FrmSeguros.frx":178BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEmpleadoras 
      Height          =   855
      Left            =   15720
      TabIndex        =   86
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "AFILIADAS"
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
      MICON           =   "FrmSeguros.frx":17D0E
      PICN            =   "FrmSeguros.frx":17D2A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCAR :"
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
      Left            =   750
      TabIndex        =   16
      Top             =   7710
      Width           =   735
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEGUROS MEDICOS"
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
      Left            =   270
      TabIndex        =   15
      Top             =   120
      Width           =   1515
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   7590
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   16815
   End
End
Attribute VB_Name = "FrmSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede

Public Sub actualizar()
  strCadena = "SELECT * FROM seguro_medico_detalle WHERE ruc='" & KEY_RUC & "' ORDER BY eps DESC,descripcion"
  Call llenarGrid(Me.HfSeguros, Me)
End Sub
Public Sub LLENA(ByVal in_seguro As String)
Dim in_distrito As String
Dim in_provincia As String
Dim in_departamento As String

strCadena = "SELECT * FROM view_seguro WHERE id_seguro='" & in_seguro & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtRUC.Text = rst("ruc_seguro")
   Me.TxtDescripcion.Text = rst("descripcion")
   Me.TxtDescuentoproducto.Text = rst("descuento_productos")
   Me.txtdireccion.Text = rst("direccion")
   Me.txtfactoraseguradora.Text = rst("factor")
   Me.txteps.Text = rst("eps")
   Me.txtvalorconsulta.Text = rst("valor_consulta")
   Me.txtdetalle.Text = rst("detalle")
   Me.txtdominio.Text = rst("dominio")
   Me.txtmail.Text = rst("mail")
   Me.txtcodaseguradora.Text = rst("cod_aseguradora")
   Me.dtcVinculado.BoundText = rst("id_vinculado")
   Me.txtcama.Text = rst("cama")
   Me.txtuci.Text = rst("uci")
   Me.txtucin.Text = rst("ucin")
   Me.txtfarmacia.Text = rst("farmacia")
   
   If rst("activo") = "si" Then
      Me.chkhabilitado.Value = 1
   Else
      Me.chkhabilitado.Value = 0
   End If
   
   If rst("id_servicios") = "si" Then
      Me.chk_pago_servicio.Value = 1
   Else
      Me.chk_pago_servicio.Value = 0
   End If
   
   
   If rst("configuracion_especial") = "si" Then
      Me.chkconfiguracion.Value = 1
   Else
      Me.chkconfiguracion.Value = 0
   End If
   
   If IsNull(rst("id_distrito")) = True Then
        in_distrito = 0
   Else
        in_distrito = rst("id_distrito")
   End If
   If IsNull(rst("id_provincia")) = True Then
        in_provincia = 0
   Else
        in_provincia = rst("id_provincia")
   End If
   If IsNull(rst("id_departamento")) = True Then
        in_departamento = 0
   Else
        in_departamento = rst("id_departamento")
   End If
   
   If Val(in_distrito) > 0 Then
        strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE  id_provincia='" & rst("id_provincia") & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcDistrito)
        Me.DtcDistrito.BoundText = in_distrito
    End If
 
 If Val(in_provincia) > 0 Then
    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE  id_departamento='" & rst("id_departamento") & "'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcProvincia)
    Me.DtcProvincia.BoundText = in_provincia
End If

 If Val(in_departamento) > 0 Then
    strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & rst("id_departamento") & "'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcDepartamento)
    Me.DtcDepartamento.BoundText = in_departamento
End If

Call listar_telefono(Me.HfTelefono, Trim(Me.txtRUC.Text))
Me.frmdetalle.Visible = True
  
   
End If


End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
'On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 5300
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1300
        Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION ASEGURADORA" & vbTab & "V.CONSULTA" & vbTab & "FACTOR" & vbTab & "CAMA" & vbTab & "UCI" & vbTab & "UCIN" & vbTab & "FARMACIA" & vbTab & "RUC"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            If rst("activo") = "si" Then
                in_estado = "ACTIVO"
            Else
                in_estado = "INACTIVO"
            End If
             Fila = rst("id_detalle") & vbTab & rst("descripcion") & vbTab & rst("valor_consulta") & vbTab & rst("factor") & vbTab & rst("cama") & vbTab & rst("uci") & vbTab & rst("ucin") & vbTab & rst("farmacia") & Space(1) & "%" & vbTab & rst("ruc_seguro")
             Grilla.AddItem Fila
             
             If rst("eps") = "1" Then
                    For k = 0 To 8
                        Grilla.col = k
                        Grilla.Row = i
                        Grilla.CellBackColor = &H80FF&
                    Next k
             End If
             If rst("activo") = "no" Then
                    For k = 0 To 8
                        Grilla.col = k
                        Grilla.Row = i
                        Grilla.CellBackColor = &H8080FF
                    Next k
             End If
             
             rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGrid_empleadora(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
'On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.rows = 0
   
    Exit Sub

End If
   
   Grilla.rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
        Next
         cabecera = "ID" & vbTab & "RUC EMP" & vbTab & "EMPLEADORA" & vbTab & "DESCTO (%)" & vbTab & "FECHA REG"
         Grilla.AddItem cabecera
         For k = 1 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        Me.lblempresa_emp.Caption = Me.lblempresa_emp.Caption & Space(2) & "::::::::[" & Trim(rst.RecordCount) & "] :::::::::"
        For i = 1 To rst.RecordCount
            
             Fila = rst("id_seguro_empleadora") & vbTab & rst("ruc_empresa") & vbTab & rst("nombre_completo") & vbTab & rst("descuento") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY")
             Grilla.AddItem Fila
            rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub ClbAcciones_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub cmd_editar_emp_Click()
Me.txt_id_seguro_empleadora.Text = Val(Me.HfEmpleadoras.TextMatrix(Me.HfEmpleadoras.Row, 0))
Me.txtruc_emp.Text = Me.HfEmpleadoras.TextMatrix(Me.HfEmpleadoras.Row, 1)
Me.txtrazon_emp.Text = Me.HfEmpleadoras.TextMatrix(Me.HfEmpleadoras.Row, 2)
Me.txtdescuento_emp.Text = Format(Me.HfEmpleadoras.TextMatrix(Me.HfEmpleadoras.Row, 3), "#,##0.00")
Me.frmdescuento.Visible = True
Call Resalta(Me.txtdescuento_emp)
 
End Sub

Private Sub cmd_salir_emp_Click()
Me.frmempleadoras.Visible = False
End Sub

Private Sub CmdAgregar_Click()
Me.frmtelefonos.Visible = True
Me.TxtFono.Text = ""
Me.txtreferencia.Text = ""
Call Resalta(Me.TxtFono)
End Sub

Private Sub cmdagregartelefono_Click()
Dim cPersona As Double
Dim key_reg As String
strCadena = "SELECT * FROM persona_telefono WHERE telefono='" & Trim(Me.TxtFono.Text) & "' and dni='" & Trim(Me.txtRUC.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 5 Then
    MsgBox "Telefono ya Registrado", vbInformation, "Mensaje para el Usuario"
    Set rst = Nothing
    Call Resalta(Me.TxtFono)
    Exit Sub
Else
    If (Me.txtRUC.Text) <> "" Then
        
        strCadena = "INSERT INTO persona_telefono (dni,telefono,id_cargo,referencia)VALUES ('" & Trim(Me.txtRUC.Text) & "','" & Trim(Me.TxtFono.Text) & "','" & Me.DtcArea.BoundText & "','" & Trim(Me.txtreferencia.Text) & "')"
        CnBd.Execute (strCadena)
        
        
       
        
        
        
    Else
        MsgBox "Ingrese un Ruc/DNI", vbInformation, "Mensaje para el Usuario"
        Call Resalta(Me.txtRUC)
        Exit Sub
    End If
    
    Me.TxtFono.Text = ""
    Call Resalta(Me.TxtFono)
    Call listar_telefono(Me.HfTelefono, Trim(Me.txtRUC.Text))
    Me.frmtelefonos.Visible = False
End If
End Sub



Private Sub cmdCerrar_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdcerrar_factor_Click()
Me.frmdescuento.Visible = False
End Sub

Private Sub cmdEditar_Click()
Me.txtid_seguro.Text = Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 0)
      Call Me.LLENA(Trim(Me.txtid_seguro.Text))
      Exit Sub
End Sub

Private Sub cmdEliminar_Click()
Procedencia = Eliminar
      Call disabled_form(Me)
      frmsegurity.Show
      Exit Sub
End Sub

Private Sub cmdEmpleadoras_Click()
Call llenar_empleadora(Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 0))
End Sub
Public Sub llenar_empleadora(ByVal in_seguro As String)
If Val(in_seguro) > 0 Then
   Me.lblempresa_emp.Caption = Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 1)
End If
strCadena = "SELECT * FROM view_seguro_empleadora where id_seguro='" & in_seguro & "' and ruc='" & KEY_RUC & "'"
Call llenarGrid_empleadora(Me.HfEmpleadoras, Me)
Me.frmempleadoras.Visible = True
End Sub


Private Sub cmdnuevo_Click()
Procedencia = Nuevos
      Me.cmdVisualizar.Visible = False
      Me.frmdetalle.Visible = True
      Me.txtid_seguro.Text = ""
      Me.TxtDescripcion.Text = ""
      Me.TxtDescuentoproducto.Text = ""
      Me.txtdetalle.Text = ""
      Me.TxtDescuentoproducto.Text = 0
      Me.txtvalorconsulta.Text = 0
      Me.txteps.Text = 0
      Me.txtcama.Text = 0
      Me.txtuci.Text = 0
      Me.txtucin.Text = 0
      Me.txtfarmacia.Text = 0
      Me.TxtFono.Text = ""
      Me.txtRUC.Text = ""
      Me.txtdireccion.Text = ""
      Me.txtfactoraseguradora.Text = 0
      Me.HfTelefono.rows = 0
      
      Call Resalta(Me.txtRUC)
      Exit Sub
End Sub

Private Sub cmdprocesar_Click()
Dim strcodigo As String
Dim in_habilitado As String

If Trim(Me.TxtDescripcion.Text) <> "" Then
    If Me.chkhabilitado.Value = 1 Then
        in_habilitado = "si"
    Else
        in_habilitado = "no"
    End If
    
    If Me.chkvincular.Value = 1 Then
        in_vinculado = Me.dtcVinculado.BoundText
    Else
        in_vinculado = 0
    End If
    
    If Me.chkconfiguracion.Value = 1 Then
        in_especial = "si"
    Else
        in_especial = "no"
    End If
    If Me.chk_pago_servicio.Value = 1 Then
        in_servicio = "si"
    Else
        in_servicio = "no"
    End If
    
    
    If Val(Me.txtid_seguro.Text) < 1 Then
        
        
        
        strCadena = "SELECT * FROM seguro_medico_detalle WHERE ruc='" & KEY_RUC & "' ORDER BY id_detalle DESC LIMIT 0,1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.txtid_seguro.Text = formato_item(Val(rst("id_detalle")) + 1, 5)
        Else
            Me.txtid_seguro.Text = formato_item(1, 5)
        End If
        
        
        
        
        
        
        strCadena = "INSERT INTO seguro_medico_detalle(`id_detalle`,configuracion_especial,id_vinculado,`ruc_seguro`,`descripcion`,`detalle`,`fecha_registro`,`descuento_productos`,`factor`,`valor_consulta`,dominio,`eps`,activo,cod_aseguradora,cama,uci,ucin,farmacia,id_servicios,`ruc`) " & _
        " VALUES('" & Trim(Me.txtid_seguro.Text) & "','" & in_especial & "','" & in_vinculado & "','" & Trim(Me.txtRUC.Text) & "','" & Trim(Me.TxtDescripcion.Text) & "','" & Trim(Me.txtdetalle.Text) & "',CURDATE(),'" & Val(Me.TxtDescuentoproducto.Text) & "','" & Val(Me.txtfactoraseguradora.Text) & "','" & Val(Me.txtvalorconsulta.Text) & "','" & Trim(Me.txtdominio.Text) & "','" & Val(Me.txteps.Text) & "','" & in_habilitado & "','" & Trim(Me.txtcodaseguradora.Text) & "','" & Val(Me.txtcama.Text) & "','" & Val(Me.txtuci.Text) & "','" & Val(Me.txtucin.Text) & "','" & Val(Me.txtfarmacia.Text) & "','" & in_servicio & "','" & KEY_RUC & "')"
    Else
        strCadena = "UPDATE seguro_medico_detalle SET  id_servicios='" & in_servicio & "', configuracion_especial='" & in_especial & "', cama='" & Val(Me.txtcama.Text) & "',uci='" & Val(Me.txtuci.Text) & "',ucin='" & Val(Me.txtucin.Text) & "',farmacia='" & Val(Me.txtfarmacia.Text) & "',   id_vinculado='" & in_vinculado & "',activo='" & in_habilitado & "', ruc_seguro='" & Trim(Me.txtRUC.Text) & "',dominio='" & Trim(Me.txtdominio.Text) & "',descripcion='" & Trim(Me.TxtDescripcion.Text) & "',detalle='" & Trim(Me.txtdetalle.Text) & "', " & _
        " descuento_productos='" & Val(Me.TxtDescuentoproducto.Text) & "',factor='" & Val(Me.txtfactoraseguradora.Text) & "',valor_consulta='" & Val(Me.txtvalorconsulta.Text) & "',eps='" & Val(Me.txteps.Text) & "',cod_aseguradora='" & Trim(Me.txtcodaseguradora.Text) & "' WHERE id_detalle='" & Trim(Me.txtid_seguro.Text) & "' and ruc='" & KEY_RUC & "'"
        
    End If
    CnBd.Execute (strCadena)
    
    Call proceso_persona(Trim(Me.txtRUC.Text), Trim(Me.TxtDescripcion.Text), Trim(Me.txtmail.Text), Trim(Me.txtdireccion.Text))
    Me.frmdetalle.Visible = False
    Call actualizar
End If
End Sub



Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub CmdVisualizar_Click()
strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRUC.Text) & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.TxtDescripcion.Text = rst("nombre_completo")
   Me.txtdireccion.Text = rst("direccion")
   Call listar_telefono(Me.HfTelefono, Trim(Me.txtRUC.Text))
Else
   Me.TxtDescripcion.Text = ""
   Me.txtdireccion.Text = ""
   Me.HfTelefono.rows = 0
End If
End Sub

Private Sub DtcDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        strCadena = "SELECT id_provincia FROM distrito WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            
            Me.DtcProvincia.Visible = True
            strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_provincia='" & rstTemporal("id_provincia") & "'"
            Call ConfiguraRst(strCadena)
            Call LlenaDataCombo(Me.DtcProvincia)
            Set rst = Nothing
            Me.DtcProvincia.Enabled = False
        End If
        
     
     
     
    
        
        
    End If
End If
End Sub

Private Sub DtcProvincia_Change()
If Me.DtcProvincia.BoundText <> "" Then
    strCadena = "SELECT * FROM provincia WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' "
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
        Me.lbldepartamento.Visible = True
        Me.DtcDepartamento.Visible = True
        strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & rstTemporal("id_departamento") & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcDepartamento)
       
        Me.DtcDepartamento.Enabled = True
    End If
    Set rstTemporal = Nothing
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
strCadena = "SELECT id_cargo as Codigo, descripcion as Descripcion FROM persona_cargos WHERE ruc='no' ORDER BY descripcion"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcArea)

strCadena = "SELECT id_detalle as Codigo,descripcion as Descripcion FROM seguro_medico_detalle WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcVinculado)



Call actualizar

Me.cmdEditar.Enabled = False
Me.cmdelete.Enabled = False


End Sub



Private Sub HfEmpleadoras_SelChange()
If Val(Me.HfEmpleadoras.TextMatrix(Me.HfEmpleadoras.Row, 0)) > 0 Then
   Me.cmd_editar_emp.Enabled = True
   Me.cmd_eliminar_emp.Enabled = True
Else
   Me.cmd_editar_emp.Enabled = False
   Me.cmd_eliminar_emp.Enabled = False
End If
End Sub

Private Sub HfSeguros_KeyPress(KeyAscii As Integer)
If Val(Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 0)) > 0 Then
    If KeyAscii = 13 Then
        If frmretencion.Procedencia = Selecionar Then
            frmretencion.Procedencia = Neutro
            Call frmretencion.load_seguro(Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 0))
            Unload Me
            Exit Sub
        End If
    End If
End If


End Sub

Private Sub HfSeguros_SelChange()
If Val(Me.HfSeguros.TextMatrix(Me.HfSeguros.Row, 0)) > 0 Then
    Me.cmdEditar.Enabled = True
    Me.cmdelete.Enabled = True
    
Else
    Me.cmdEditar.Enabled = False
    Me.cmdelete.Enabled = False
End If
End Sub

Private Sub s_Click()

End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM seguro_medico_detalle WHERE descripcion LIKE '%" & Trim(Me.TXTBUSCAR.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY  eps DESC,descripcion"
    Call llenarGrid(Me.HfSeguros, Me)
    
End If
End Sub

Private Sub TxtDistrito_Change()
If Trim(Me.TxtDistrito.Text) <> "" Then
    strCadena = "SELECT id_distrito as Codigo,CONCAT(d.descripcion,' -- > ',p.descripcion) as Descripcion FROM distrito d, provincia p WHERE  p.id_provincia=d.id_provincia and   d.descripcion LIKE '%" & Trim(Me.TxtDistrito.Text) & "%'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    
End If
End Sub

Private Sub TxtDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        Me.DtcDistrito.SetFocus
    End If
End If
End Sub

Private Sub TxtRuc_Change()
If Len(Trim(Me.txtRUC.Text)) >= 8 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRUC.Text) & "' LIMIT 1"
    Call ConfiguraRstAux(strCadena)
    If rstAux.RecordCount > 0 Then
        Me.cmdVisualizar.Visible = True
    Else
        Me.cmdVisualizar.Visible = False
    End If
End If
End Sub

Private Sub txtvincular_Change()

strCadena = "SELECT id_detalle as Codigo,descripcion as Descripcion FROM seguro_medico_detalle WHERE descripcion LIKE '%" & Trim(Me.txtvincular.Text) & "%' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcVinculado)
End Sub

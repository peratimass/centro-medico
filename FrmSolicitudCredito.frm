VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmSolicitudCredito 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   18945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle_credito 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SOLICITUD CREDITO"
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
      Height          =   6495
      Left            =   3600
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   11175
      Begin VB.TextBox txtObservacion 
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
         Height          =   1455
         Left            =   6600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   3120
         Width           =   4335
      End
      Begin VB.Frame frmAceptacion 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   1200
         TabIndex        =   25
         Top             =   5160
         Visible         =   0   'False
         Width           =   5175
         Begin VB.OptionButton opt_rechazar 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "RECHAZAR CREDITO"
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
            Height          =   300
            Left            =   360
            TabIndex        =   27
            Top             =   600
            Width           =   3135
         End
         Begin VB.OptionButton Opt_aceptar 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            Caption         =   "ACEPTAR CREDITO"
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
            Height          =   300
            Left            =   360
            TabIndex        =   26
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.TextBox txtNumero 
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
         TabIndex        =   14
         Top             =   960
         Width           =   1455
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
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   780
         Left            =   8880
         TabIndex        =   16
         Top             =   5520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         MICON           =   "FrmSolicitudCredito.frx":0000
         PICN            =   "FrmSolicitudCredito.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSalir_detalle 
         Height          =   780
         Left            =   9960
         TabIndex        =   17
         Top             =   5520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         MICON           =   "FrmSolicitudCredito.frx":3664
         PICN            =   "FrmSolicitudCredito.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI/RUC :"
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
         TabIndex        =   34
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lblruc 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   33
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblfechasolicitud 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   31
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.SOLICITUD :"
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
         TabIndex        =   30
         Top             =   4320
         Width           =   915
      End
      Begin VB.Label lblid_venta 
         BackColor       =   &H00C0C0C0&
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
         Left            =   3480
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblidsolicitud 
         BackColor       =   &H00FFFFFF&
         Caption         =   "123456"
         BeginProperty Font 
            Name            =   "3 of 9 Barcode"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblsolicitante 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   24
         Top             =   3840
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITANTE:"
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
         Left            =   255
         TabIndex        =   23
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label lblMontoCredito 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   22
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblvendedor 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   21
         Top             =   3360
         Width           =   5055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDEDOR."
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
         TabIndex        =   20
         Top             =   3480
         Width           =   795
      End
      Begin VB.Label lblcelular 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   19
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lbldireccion 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   18
         Top             =   2160
         Width           =   8415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.CREDITO :"
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
         Left            =   315
         TabIndex        =   15
         Top             =   3000
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROFORMA :"
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
         Left            =   300
         TabIndex        =   12
         Top             =   1035
         Width           =   855
      End
      Begin VB.Label lblcliente 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1800
         Width           =   8415
      End
   End
   Begin VB.TextBox TxtProveedor 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   345
      Left            =   8280
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "FrmSolicitudCredito.frx":3A70
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
      Height          =   345
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   63700993
      CurrentDate     =   43141
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   17655
      _ExtentX        =   31141
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
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   900
      Left            =   17880
      TabIndex        =   4
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      MICON           =   "FrmSolicitudCredito.frx":3A8C
      PICN            =   "FrmSolicitudCredito.frx":3AA8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEditable 
      Height          =   900
      Left            =   17880
      TabIndex        =   5
      Top             =   2145
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "CONFIRMAR"
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
      MICON           =   "FrmSolicitudCredito.frx":3EFA
      PICN            =   "FrmSolicitudCredito.frx":3F16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   900
      Left            =   17880
      TabIndex        =   6
      Top             =   4035
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "CERRAR"
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
      MICON           =   "FrmSolicitudCredito.frx":4230
      PICN            =   "FrmSolicitudCredito.frx":424C
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
      Height          =   345
      Left            =   6600
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   63700993
      CurrentDate     =   43141
   End
   Begin VitekeySoft.ChameleonBtn cmdreporte 
      Height          =   855
      Left            =   17880
      TabIndex        =   32
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "E.CUENTA"
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
      MICON           =   "FrmSolicitudCredito.frx":7273
      PICN            =   "FrmSolicitudCredito.frx":728F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOLICITUD CREDITO"
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
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   45
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE :"
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
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   17655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8520
      Left            =   0
      Top             =   0
      Width           =   18945
   End
End
Attribute VB_Name = "FrmSolicitudCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdbuscar_Click()

strCadena = "SELECT * FROM view_solicitud_credito_ultimate WHERE fecha_solicitud>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgDetalle)


End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me
End Sub

Private Sub cmdEditable_Click()
If Val(Me.HfgDetalle.Rows) > 0 Then
   If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
       Call get_solicitud(Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)))
   End If
End If
End Sub

Private Sub cmdNuevo_Click()
Call nuevo
End Sub
Private Sub llenar()

strCadena = "SELECT * FROM view_solicitud_credito_ultimate WHERE ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgDetalle)



End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
  
 
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 2500
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1400
          Next
         cabecera = "SOLICITUD" & vbTab & "F.SOLICITUD" & vbTab & "COMPROBANTE" & vbTab & "MONTO" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "SOLICITADO POR" & vbTab & "CONFIRMADO POR" & vbTab & "F.CONFIRMACION" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Select Case rst("estado")
                    Case "01"
                        in_estado = "EN EVALUACION"
                    Case "02"
                        in_estado = "ACEPTADO"
                    Case "03"
                        in_estado = "RECHAZADO"
             End Select
        
             Fila = Format(rst("id_solicitud"), "000000") & vbTab & Format(rst("fecha_solicitud"), "dd-mm-YYYY") & vbTab & rst("documento") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & Mid(rst("enviado"), 1, 40) & vbTab & rst("confirma") & vbTab & rst("fecha_aceptacion") & vbTab & in_estado
             Grilla.AddItem Fila
                            If rst("estado") = "01" Then
                                For k = 7 To 9
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &HC0C0FF
                                Next k
                            End If
                            If rst("estado") = "02" Then
                                For k = 7 To 9
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &HC0FFC0
                                Next k
                            End If
                            If rst("estado") = "03" Then
                                For k = 7 To 9
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &H80C0FF
                                Next k
                            End If
                            
            
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub nuevo()

Me.frmdetalle_credito.Visible = True
Me.lblidsolicitud.Caption = 0
Me.txtSerie.Text = ""
Me.txtNumero.Text = ""
Me.lblcelular.Caption = ""
Me.lblcliente.Caption = ""
Me.lbldireccion.Caption = ""
Me.lblsolicitante.Caption = KEY_VENDEDOR
Me.lblMontoCredito.Caption = 0
Me.lblvendedor.Caption = ""
Me.lblid_venta.Caption = ""
Me.lblfechasolicitud.Caption = ""
Me.txtObservacion.Text = ""
Me.cmdprocesar.Enabled = False
Call Resalta(Me.txtSerie)

End Sub
Private Sub get_solicitud(ByVal in_solicitud As String)
strCadena = "SELECT * FROM solicitud_credito WHERE id_solicitud='" & Val(in_solicitud) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Me.frmAceptacion.Visible = True
   If rstK("estado") = "02" Then
      Me.Opt_aceptar.Value = True
   Else
      Me.opt_rechazar.Value = True
   End If
   
   Me.lblidsolicitud.Caption = Format(rstK("id_solicitud"), "000000")
   Me.lblfechasolicitud.Caption = Format(rstK("fecha_solicitud"), "dd-mm-YYYY")
   Me.txtObservacion.Text = rstK("observacion")
  
   
   strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rstK("id_proforma") & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Call get_proforma(rst("serie"), rst("numero"))
      
   End If
   Me.frmdetalle_credito.Visible = True
End If

End Sub


Private Sub get_proforma(ByVal in_serie As String, ByVal in_numero As String)

strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='0099' and serie='" & Format(in_serie, "000") & "' and numero='" & Format(in_numero, "000000") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.lblruc.Caption = rst("id_cliente")
   Me.txtSerie.Text = rst("serie")
   Me.txtNumero.Text = rst("numero")
   Me.lblid_venta.Caption = rst("id_venta")
   Me.lblcliente.Caption = rst("ncliente")
   Me.lbldireccion.Caption = rst("direccion")
   Me.lblcelular.Caption = get_telefono(rst("id_cliente"))
   Me.lblMontoCredito.Caption = Format(rst("total"), "#,##0.00")
   Me.lblvendedor.Caption = get_persona(rst("dni_save"))
   Me.lblsolicitante.Caption = KEY_VENDEDOR
   Me.cmdprocesar.Enabled = True
   Me.cmdreporte.Enabled = True
End If
End Sub

Private Sub cmdprocesar_Click()
If Val(Me.lblidsolicitud.Caption) > 0 Then
   If Me.Opt_aceptar.Value = True Then
      in_estado = "02"
   Else
      in_estado = "03"
   End If
Else
    in_estado = "01"
End If

strCadena = "call put_solicitud_credito('" & Val(Me.lblidsolicitud.Caption) & "','" & Val(Me.lblid_venta.Caption) & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & in_estado & "','" & UCase(Me.txtObservacion.Text) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


MsgBox "Procesado con Exito.", vbInformation
Unload Me
Call llenar
End Sub

Private Sub cmdreporte_Click()
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant

Dim in_ruc As String
If Len(Trim(Me.lblruc.Caption)) = 8 Then
   in_ruc = "10" & Trim(Me.lblruc.Caption)
   in_ruc = DigitoVerificadorRUC(Trim(Me.lblruc.Caption))
Else
    in_ruc = Trim(Me.lblruc.Caption)
End If

arr(0, 1) = "in_ruc"
arr(0, 2) = in_ruc
param = arr()

  strCadena = "SELECT id_venta,(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_doc " & _
  " FROM view_listado_comprobante_vitekey WHERE anulado='no' and  id_cliente LIKE '%" & Trim(Me.lblruc.Caption) & "%' AND ruc='" & KEY_RUC & "'"


Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
           If rst("id_doc") = "0007" Then
                n_saldo = rst("saldo") * -1
           Else
                n_saldo = rst("saldo")
           End If
        strCadena = "UPDATE movimiento_venta SET saldo='" & n_saldo & "' WHERE id_venta='" & rst("id_venta") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
    Next i
End If


  strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,hora,numero,comprobante,id_cliente,ncliente,direccion,celular,total, saldo  ,anulado,id_moneda,ruc,tc,id_alm,id_doc,referencia,descripcion,simbolo FROM view_historial_pagos_cobrar WHERE anulado='no' and  saldo<>0 and id_cliente LIKE '%" & Trim(Me.lblruc.Caption) & "%' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza_letras", param, App.Path + "\Reportes\")

End Sub

Private Sub cmdSalir_detalle_Click()
Me.frmdetalle_credito.Visible = False
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Call llenar
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtNumero.Text = Format(Me.txtNumero.Text, "000000")
    Call get_proforma(Trim(Me.txtSerie.Text), Trim(Me.txtNumero.Text))
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtSerie.Text = Format(Me.txtSerie.Text, "000")
   Call Resalta(Me.txtNumero)
End If
End Sub

VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmDetallePersona 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de Clientes"
   ClientHeight    =   9240
   ClientLeft      =   540
   ClientTop       =   315
   ClientWidth     =   20145
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frmUbigeo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2325
      Left            =   5085
      TabIndex        =   283
      Top             =   6885
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox txtBuscaUbigeo 
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
         Height          =   315
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   286
         Top             =   120
         Width           =   2535
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfUbigeo 
         Height          =   1815
         Left            =   120
         TabIndex        =   284
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3201
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   8421376
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
      Begin VB.Image Image2 
         Height          =   240
         Left            =   7800
         Picture         =   "FrmDetalleCliente.frx":08CA
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblDepartamento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR UBIGEO:"
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
         Index           =   0
         Left            =   150
         TabIndex        =   285
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.TextBox txtZona 
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
      Height          =   300
      Left            =   8805
      MaxLength       =   80
      TabIndex        =   255
      Top             =   8540
      Width           =   615
   End
   Begin VB.CheckBox chk_extranjeria 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "C.EXTRANJERIA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   234
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox chk_igv 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "A.IGV"
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
      Height          =   250
      Left            =   7200
      TabIndex        =   226
      Top             =   4155
      Width           =   930
   End
   Begin VB.Frame frmTransporte 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4335
      Left            =   11445
      TabIndex        =   186
      Top             =   4455
      Visible         =   0   'False
      Width           =   8655
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   120
         TabIndex        =   187
         Top             =   120
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7223
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "MARCAS / PLACAS"
         TabPicture(0)   =   "FrmDetalleCliente.frx":376E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdEliminaMarca"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdNuevoMarca"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "HfMarcas"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "frmmarca"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "CONDUCTORES"
         TabPicture(1)   =   "FrmDetalleCliente.frx":378A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmconductores"
         Tab(1).Control(1)=   "HfChofer"
         Tab(1).Control(2)=   "cmdNuevoChofer"
         Tab(1).Control(3)=   "cmdEliminarchofer"
         Tab(1).ControlCount=   4
         Begin VB.Frame frmconductores 
            Caption         =   "DETALLE MARCA"
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
            Height          =   2895
            Left            =   -74760
            TabIndex        =   201
            Top             =   480
            Visible         =   0   'False
            Width           =   6495
            Begin VB.TextBox lblchofer 
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
               Height          =   315
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   241
               Top             =   960
               Width           =   4815
            End
            Begin VB.TextBox txtchofer 
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
               Height          =   315
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   203
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtlicenciaTransporte 
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
               Height          =   315
               Left            =   1560
               MaxLength       =   80
               TabIndex        =   202
               Top             =   1560
               Width           =   1815
            End
            Begin VitekeySoft.ChameleonBtn cmdprocesar_chofer 
               Height          =   300
               Left            =   1560
               TabIndex        =   204
               Top             =   2160
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
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
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmDetalleCliente.frx":37A6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdCerrar_chofer 
               Height          =   300
               Left            =   2520
               TabIndex        =   205
               Top             =   2160
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
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
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmDetalleCliente.frx":37C2
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
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CHOFER :"
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
               Index           =   2
               Left            =   735
               TabIndex        =   267
               Top             =   1080
               Width           =   645
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LICENCIA :"
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
               Index           =   1
               Left            =   645
               TabIndex        =   266
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DNI :"
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
               Index           =   0
               Left            =   1035
               TabIndex        =   206
               Top             =   480
               Width           =   345
            End
         End
         Begin VB.Frame frmmarca 
            Caption         =   "DETALLE MARCA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3495
            Left            =   240
            TabIndex        =   194
            Top             =   480
            Visible         =   0   'False
            Width           =   6975
            Begin VB.TextBox TxtColor 
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
               Left            =   1275
               TabIndex        =   239
               Top             =   2760
               Width           =   3135
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
               Left            =   1275
               TabIndex        =   238
               Top             =   2280
               Width           =   3135
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
               Left            =   1275
               TabIndex        =   236
               Top             =   1800
               Width           =   3135
            End
            Begin VB.TextBox TxtCertificado 
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
               Height          =   315
               Left            =   1275
               MaxLength       =   80
               TabIndex        =   210
               Top             =   1320
               Width           =   3135
            End
            Begin VB.TextBox TxtIdMarca 
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
               Height          =   315
               Left            =   4560
               MaxLength       =   80
               TabIndex        =   207
               Top             =   360
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox txtPlaca 
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
               Height          =   315
               Left            =   1275
               MaxLength       =   80
               TabIndex        =   198
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtMarca 
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
               Height          =   315
               Left            =   1275
               MaxLength       =   80
               TabIndex        =   197
               Top             =   840
               Width           =   2055
            End
            Begin VitekeySoft.ChameleonBtn cmdprocesarMarca 
               Height          =   300
               Left            =   4800
               TabIndex        =   199
               Top             =   2760
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
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
               BCOL            =   8421631
               BCOLO           =   8421631
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmDetalleCliente.frx":37DE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdCerrarMarca 
               Height          =   300
               Left            =   5760
               TabIndex        =   200
               Top             =   2760
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   529
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
               BCOL            =   8421631
               BCOLO           =   8421631
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmDetalleCliente.frx":37FA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "COLOR :"
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
               Left            =   525
               TabIndex        =   240
               Top             =   2880
               Width           =   540
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N� MOTOR  :"
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
               TabIndex        =   237
               Top             =   2400
               Width           =   825
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N� SERIE  :"
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
               Left            =   405
               TabIndex        =   235
               Top             =   1800
               Width           =   660
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CERTIFICADO :"
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
               Left            =   60
               TabIndex        =   211
               Top             =   1320
               Width           =   1005
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MARCA  :"
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
               Left            =   390
               TabIndex        =   196
               Top             =   840
               Width           =   675
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PLACA :"
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
               Left            =   510
               TabIndex        =   195
               Top             =   360
               Width           =   555
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfMarcas 
            Height          =   3375
            Left            =   240
            TabIndex        =   188
            Top             =   480
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   5953
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   8388608
            FocusRect       =   0
            GridLines       =   2
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
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
         Begin VitekeySoft.ChameleonBtn cmdNuevoMarca 
            Height          =   300
            Left            =   7320
            TabIndex        =   189
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":3816
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdEliminaMarca 
            Height          =   300
            Left            =   7320
            TabIndex        =   190
            Top             =   885
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            BTYPE           =   5
            TX              =   "ELIMINAR"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":3832
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfChofer 
            Height          =   2895
            Left            =   -74760
            TabIndex        =   191
            Top             =   480
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5106
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   8388608
            FocusRect       =   0
            GridLines       =   2
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
         Begin VitekeySoft.ChameleonBtn cmdNuevoChofer 
            Height          =   300
            Left            =   -68160
            TabIndex        =   192
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
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
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":384E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdEliminarchofer 
            Height          =   300
            Left            =   -68160
            TabIndex        =   193
            Top             =   1005
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            BTYPE           =   5
            TX              =   "ELIMINAR"
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
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":386A
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
   End
   Begin VB.CheckBox chkPais 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "PAIS :"
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
      Height          =   250
      Left            =   210
      TabIndex        =   157
      Top             =   7850
      Width           =   1095
   End
   Begin VB.CommandButton cmdConSUNAT 
      BackColor       =   &H008080FF&
      Caption         =   "RENIEC"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   153
      Tag             =   "0"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtsectorista 
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
      Height          =   300
      Left            =   8805
      MaxLength       =   80
      TabIndex        =   109
      Top             =   8880
      Width           =   615
   End
   Begin VB.Frame frmdireccion 
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1560
      TabIndex        =   92
      Top             =   1995
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdcerrardireccion 
         Height          =   270
         Left            =   5160
         Picture         =   "FrmDetalleCliente.frx":3886
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   120
         Width           =   270
      End
      Begin VB.TextBox txtdireccion3 
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
         Height          =   315
         Left            =   4140
         MaxLength       =   80
         TabIndex        =   95
         Top             =   1335
         Width           =   1095
      End
      Begin VB.TextBox txtnuevadireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   120
         MaxLength       =   150
         TabIndex        =   94
         Top             =   140
         Width           =   4935
      End
      Begin VitekeySoft.ChameleonBtn cmdagregar_direccion 
         Height          =   300
         Left            =   4140
         TabIndex        =   93
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "AGREGAR"
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":672A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcDistrito2 
         Height          =   315
         Left            =   1320
         TabIndex        =   96
         Top             =   1335
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcDepartamento2 
         Height          =   315
         Left            =   1320
         TabIndex        =   97
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcProvincia2 
         Height          =   315
         Left            =   1320
         TabIndex        =   98
         Top             =   900
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin VB.Label lblCodigoUbigeoSunat2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   1320
         TabIndex        =   288
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRITO :"
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
         Index           =   4
         Left            =   495
         TabIndex        =   273
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA  :"
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
         Index           =   3
         Left            =   315
         TabIndex        =   272
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTAMENTO :"
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
         Index           =   2
         Left            =   -15
         TabIndex        =   271
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_id_direccion 
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
      Left            =   7200
      TabIndex        =   91
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo_direccion 
      Height          =   300
      Left            =   7155
      TabIndex        =   88
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
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
      BCOL            =   8438015
      BCOLO           =   8438015
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleCliente.frx":6746
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtAnio 
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
      Left            =   3480
      MaxLength       =   11
      TabIndex        =   84
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CheckBox chkDescuento 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "(%DESC)"
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
      Height          =   250
      Left            =   3650
      TabIndex        =   83
      Top             =   4155
      Width           =   855
   End
   Begin VB.TextBox TxtDescuento 
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
      Left            =   4560
      MaxLength       =   11
      TabIndex        =   82
      Text            =   "0.00"
      Top             =   4140
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "SUBIR IMAGEN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   3285
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "ALBUM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3610
      Width           =   2535
   End
   Begin VB.CommandButton cmdHuella 
      BackColor       =   &H008080FF&
      Caption         =   "CAPTURAR HUELLAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3950
      Width           =   2535
   End
   Begin VB.TextBox txtRUC 
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
      Height          =   285
      Left            =   3120
      TabIndex        =   49
      Top             =   165
      Width           =   1575
   End
   Begin VB.Frame div_verifica 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   6480
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "VISUALIZAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
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
         Height          =   720
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1590
      End
   End
   Begin VB.TextBox txtMaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   42
      Top             =   900
      Width           =   3135
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   41
      Top             =   1280
      Width           =   3135
   End
   Begin VB.TextBox txtPaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   40
      Top             =   530
      Width           =   3135
   End
   Begin VB.TextBox TxtDistrito 
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
      Height          =   315
      Left            =   4260
      MaxLength       =   80
      TabIndex        =   36
      Top             =   8880
      Width           =   735
   End
   Begin VB.TextBox txtdia 
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   2655
      Left            =   9820
      TabIndex        =   15
      Top             =   4560
      Width           =   1600
      Begin VB.CheckBox chkhabilitado 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "HABILITADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   156
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox ChkAlmacen 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SERVICES"
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
         TabIndex        =   56
         ToolTipText     =   "PERSONAL BAJO RESPONSABILIDAD DE BIENES"
         Top             =   2100
         Width           =   1200
      End
      Begin VB.CheckBox ChkCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTE"
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
         TabIndex        =   55
         Top             =   240
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CheckBox ChkProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PROVEEDOR"
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
         TabIndex        =   54
         Top             =   520
         Width           =   1200
      End
      Begin VB.CheckBox Chkcontable 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "VENDEDOR"
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
         TabIndex        =   53
         Top             =   820
         Width           =   1200
      End
      Begin VB.CheckBox ChkTransporte 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TRANSPORTE"
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
         TabIndex        =   52
         Top             =   1125
         Width           =   1200
      End
      Begin VB.CheckBox ChkPersonal 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PERSONAL"
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
         TabIndex        =   51
         Top             =   1450
         Width           =   1200
      End
      Begin VB.CheckBox ChkAuspiciador 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AUSPICIA"
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
         TabIndex        =   50
         Top             =   1770
         Width           =   1200
      End
      Begin VitekeySoft.ChameleonBtn cmdTransporte 
         Height          =   195
         Left            =   1350
         TabIndex        =   185
         Top             =   1125
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   344
         BTYPE           =   5
         TX              =   "..."
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":6762
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
   Begin VB.CheckBox ChkRetencion 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "A.RETEN"
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
      Height          =   250
      Left            =   6240
      TabIndex        =   13
      Top             =   4155
      Width           =   930
   End
   Begin VB.CheckBox ChkPercepcion 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "A.PERCEP"
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
      Height          =   250
      Left            =   5160
      TabIndex        =   12
      Top             =   4155
      Width           =   1050
   End
   Begin VB.CommandButton CmdFoto 
      BackColor       =   &H008080FF&
      Caption         =   "CAPTURAR IMAGEN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2960
      Width           =   2535
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2160
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Txttelefono1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   5
      Top             =   3600
      Width           =   5535
   End
   Begin VB.TextBox TxtTelefono2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2145
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtDireccion1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      MaxLength       =   150
      TabIndex        =   0
      Top             =   2055
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo Dtcmes 
      Height          =   330
      Left            =   2160
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
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
   Begin InetCtlsObjects.Inet inetConecta 
      Left            =   10920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser wbrInfo 
      Height          =   9045
      Left            =   11400
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   15954
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin TabDlg.SSTab SstKardex 
      Height          =   3285
      Left            =   165
      TabIndex        =   20
      Top             =   4515
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5794
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   7
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CREDITO/SEGURO"
      TabPicture(0)   =   "FrmDetalleCliente.frx":677E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk_estadoseguro"
      Tab(0).Control(1)=   "DtpFechaEmision_transporte"
      Tab(0).Control(2)=   "txtpolizaseguro_transporte"
      Tab(0).Control(3)=   "chk_seguro_transporte"
      Tab(0).Control(4)=   "chkclientemayor"
      Tab(0).Control(5)=   "txtMaximoCredito"
      Tab(0).Control(6)=   "txtRucEmpresa"
      Tab(0).Control(7)=   "OptSincredito"
      Tab(0).Control(8)=   "ChkMaximoCredito"
      Tab(0).Control(9)=   "chkEmpresa"
      Tab(0).Control(10)=   "Dtcseguroempresa"
      Tab(0).Control(11)=   "DtpFechaCaducidad_transporte"
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(13)=   "Label10"
      Tab(0).Control(14)=   "Label9"
      Tab(0).Control(15)=   "Shape6"
      Tab(0).Control(16)=   "lblEmpresa"
      Tab(0).Control(17)=   "Shape4"
      Tab(0).Control(18)=   "LblCantidad"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "PLAN DE SERVICIO"
      TabPicture(1)   =   "FrmDetalleCliente.frx":679A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmplandetalle"
      Tab(1).Control(1)=   "Hfplanservicio"
      Tab(1).Control(2)=   "cmdnuevoplan"
      Tab(1).Control(3)=   "cmdmodificarplan"
      Tab(1).Control(4)=   "cmdeliminarplan"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "TELEFONOS"
      TabPicture(2)   =   "FrmDetalleCliente.frx":67B6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtObservacion"
      Tab(2).Control(1)=   "cmdeditarTelefono"
      Tab(2).Control(2)=   "TxtLicencia"
      Tab(2).Control(3)=   "TxtFono"
      Tab(2).Control(4)=   "Command3"
      Tab(2).Control(5)=   "Command1"
      Tab(2).Control(6)=   "HfTelefonos"
      Tab(2).Control(7)=   "DtcArea"
      Tab(2).Control(8)=   "LblObservacion(0)"
      Tab(2).Control(9)=   "lblidtelefono"
      Tab(2).Control(10)=   "LblObservacion(1)"
      Tab(2).Control(11)=   "Label18"
      Tab(2).Control(12)=   "Label15"
      Tab(2).Control(13)=   "Shape7"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "ESTADO ALUMNO"
      TabPicture(3)   =   "FrmDetalleCliente.frx":67D2
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Shape3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "LblObservacion(2)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "LblObservacion(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "LblObservacion(4)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "LblObservacion(6)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "LblObservacion(7)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "LblObservacion(5)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "LblObservacion(8)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "LblObservacion(9)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "LblObservacion(10)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "LblObservacion(11)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "LblObservacion(12)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "LblObservacion(13)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "LblObservacion(14)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "LblObservacion(17)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "LblObservacion(18)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "LblObservacion(19)"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "LblObservacion(20)"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "DtcEstado"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "DtcPension"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "DtcMatricula"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "DtcPeriodo"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "DtcGrado"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "DtcTiponacimiento"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "DtcNivel"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "txtieprocedencia"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "txtpromovido"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "txttercio"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "txthabilidad"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "chkrecuperacion_si"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "chkrecuperacion_no"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "chkseguro"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "txtenfermedades"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "txtvacunas"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "txtalergias"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "txttalla"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "txtpeso"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "chk_beca"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "chk_media_beca"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "frmseguro"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).ControlCount=   40
      TabCaption(4)   =   "CTAS BANCO"
      TabPicture(4)   =   "FrmDetalleCliente.frx":67EE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdeliminarcuenta"
      Tab(4).Control(1)=   "cmdagregarcuenta"
      Tab(4).Control(2)=   "txtnumerocuenta"
      Tab(4).Control(3)=   "DtcBanco"
      Tab(4).Control(4)=   "DtcMoneda"
      Tab(4).Control(5)=   "HfCuentas"
      Tab(4).Control(6)=   "Label17"
      Tab(4).Control(7)=   "Label16"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "APODERADO"
      TabPicture(5)   =   "FrmDetalleCliente.frx":680A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdeliminar_item"
      Tab(5).Control(1)=   "chk_mama"
      Tab(5).Control(2)=   "chk_papa"
      Tab(5).Control(3)=   "txtdireccion_familia"
      Tab(5).Control(4)=   "txtDni"
      Tab(5).Control(5)=   "TxtFmaterno"
      Tab(5).Control(6)=   "cmdAgregar"
      Tab(5).Control(7)=   "TxtTelefono"
      Tab(5).Control(8)=   "TxtFpaterno"
      Tab(5).Control(9)=   "TxtFnombers"
      Tab(5).Control(10)=   "HfgFamiliares"
      Tab(5).Control(11)=   "DtcParentesco"
      Tab(5).Control(12)=   "DtcOcupacion"
      Tab(5).Control(13)=   "DtcGradoinstruccion"
      Tab(5).Control(14)=   "Label37"
      Tab(5).Control(15)=   "Label36"
      Tab(5).Control(16)=   "Label35"
      Tab(5).Control(17)=   "Label12"
      Tab(5).Control(18)=   "LblObservacion(23)"
      Tab(5).Control(19)=   "Label31"
      Tab(5).Control(20)=   "Label30"
      Tab(5).Control(21)=   "LblObservacion(25)"
      Tab(5).Control(22)=   "LblObservacion(24)"
      Tab(5).Control(23)=   "LblObservacion(26)"
      Tab(5).ControlCount=   24
      TabCaption(6)   =   "SEGURIDAD"
      TabPicture(6)   =   "FrmDetalleCliente.frx":6826
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frmpassword"
      Tab(6).Control(1)=   "frmBiblio"
      Tab(6).Control(2)=   "txtpassword"
      Tab(6).Control(3)=   "cmdActualizar_password"
      Tab(6).Control(4)=   "Label24(2)"
      Tab(6).Control(5)=   "Shape2(1)"
      Tab(6).ControlCount=   6
      Begin VB.TextBox TxtObservacion 
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
         Height          =   795
         Left            =   -68520
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   278
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Frame frmseguro 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "SEGURO DETALLE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1110
         Left            =   840
         TabIndex        =   140
         Top             =   2205
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton cmdcerrarseguro 
            BackColor       =   &H008080FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   240
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtpExpedicion 
            Height          =   300
            Left            =   1440
            TabIndex        =   145
            Top             =   765
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
            Format          =   135790593
            CurrentDate     =   42764
         End
         Begin VB.TextBox txtpoliza 
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
            Left            =   1440
            MaxLength       =   80
            TabIndex        =   144
            Top             =   465
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo DtcSeguro 
            Height          =   315
            Left            =   120
            TabIndex        =   141
            Top             =   120
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin MSComCtl2.DTPicker DtpExpiracion 
            Height          =   300
            Left            =   3000
            TabIndex        =   147
            Top             =   765
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
            Format          =   135790593
            CurrentDate     =   42764
         End
         Begin VB.Label LblObservacion 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA REGISTRO :"
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
            Index           =   16
            Left            =   135
            TabIndex        =   143
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label LblObservacion 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "POLIZA :"
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
            Index           =   15
            Left            =   765
            TabIndex        =   142
            Top             =   480
            Width           =   585
         End
      End
      Begin VB.CheckBox chk_media_beca 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "MEDIA BECA"
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
         TabIndex        =   230
         ToolTipText     =   "PERSONAL BAJO RESPONSABILIDAD DE BIENES"
         Top             =   3060
         Width           =   3240
      End
      Begin VB.CheckBox chk_beca 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "BECA INTEGRAL"
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
         Height          =   200
         Left            =   2160
         TabIndex        =   229
         ToolTipText     =   "PERSONAL BAJO RESPONSABILIDAD DE BIENES"
         Top             =   2840
         Width           =   3240
      End
      Begin VB.Frame frmpassword 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   615
         Left            =   -69840
         TabIndex        =   220
         Top             =   1920
         Visible         =   0   'False
         Width           =   4250
         Begin VB.TextBox txtpassword_confirmar 
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
            IMEMode         =   3  'DISABLE
            Left            =   1140
            MaxLength       =   100
            PasswordChar    =   "*"
            TabIndex        =   222
            Top             =   120
            Width           =   1695
         End
         Begin VitekeySoft.ChameleonBtn cmdconfirmar_pass 
            Height          =   315
            Left            =   2940
            TabIndex        =   223
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "ENV. PASSWORD"
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
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":6842
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONFIRMAR :"
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
            Index           =   3
            Left            =   0
            TabIndex        =   221
            Top             =   120
            Width           =   945
         End
      End
      Begin VB.Frame frmBiblio 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   2175
         Left            =   -74760
         TabIndex        =   215
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         Begin MSDataListLib.DataCombo DtcFacultad 
            Height          =   330
            Left            =   1200
            TabIndex        =   218
            Top             =   480
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSDataListLib.DataCombo DtcCiclo 
            Height          =   330
            Left            =   1200
            TabIndex        =   219
            Top             =   1080
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSDataListLib.DataCombo DtcTipoAcceso 
            Height          =   330
            Left            =   1200
            TabIndex        =   224
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
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
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO ACCESO :"
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
            Index           =   4
            Left            =   120
            TabIndex        =   225
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CICLO :"
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
            Index           =   1
            Left            =   570
            TabIndex        =   217
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FACULTAD :"
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
            Index           =   0
            Left            =   240
            TabIndex        =   216
            Top             =   480
            Width           =   825
         End
      End
      Begin VB.TextBox txtpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   -68720
         MaxLength       =   100
         PasswordChar    =   "*"
         TabIndex        =   213
         Top             =   1440
         Width           =   1695
      End
      Begin VitekeySoft.ChameleonBtn cmdeditarTelefono 
         Height          =   495
         Left            =   -70560
         TabIndex        =   183
         Top             =   495
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "FrmDetalleCliente.frx":685E
         PICN            =   "FrmDetalleCliente.frx":687A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton cmdeliminar_item 
         BackColor       =   &H008080FF&
         Caption         =   "ELIMINAR ITEM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   -67440
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CheckBox chk_estadoseguro 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SEGURO ACTIVO"
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
         Height          =   255
         Left            =   -69480
         TabIndex        =   180
         Top             =   2655
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFechaEmision_transporte 
         Height          =   345
         Left            =   -69480
         TabIndex        =   178
         Top             =   1860
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   135790593
         CurrentDate     =   42868
      End
      Begin VB.TextBox txtpolizaseguro_transporte 
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
         Left            =   -69480
         MaxLength       =   80
         TabIndex        =   177
         Top             =   1500
         Width           =   1935
      End
      Begin VB.CheckBox chk_seguro_transporte 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AFILIADO A UN SEGURO"
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
         Height          =   250
         Left            =   -70680
         TabIndex        =   172
         Top             =   660
         Width           =   2175
      End
      Begin VB.Frame frmplandetalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATOS DEL SERVICIO"
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
         Height          =   2865
         Left            =   -74880
         TabIndex        =   163
         Top             =   350
         Visible         =   0   'False
         Width           =   8415
         Begin VB.CheckBox chk_activacion_corte 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "ACTIVACION CORTE"
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
            Height          =   250
            Left            =   5400
            TabIndex        =   270
            Top             =   480
            Width           =   2895
         End
         Begin VB.CheckBox chk_ultimo_dia_mes 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "ULT DIA MES"
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
            Height          =   220
            Left            =   3000
            TabIndex        =   265
            Top             =   2480
            Width           =   1935
         End
         Begin VB.TextBox txtUltimodia_pago 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   4320
            MaxLength       =   80
            TabIndex        =   263
            Top             =   2200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox chk_3meses 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "PAGO 3 MESES"
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
            Height          =   220
            Left            =   1440
            TabIndex        =   246
            Top             =   2190
            Width           =   1455
         End
         Begin VB.CheckBox chk_6meses 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "PAGO 6 MESES"
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
            Height          =   220
            Left            =   1440
            TabIndex        =   245
            Top             =   1935
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DtpFechaSuscripcion 
            Height          =   300
            Left            =   3000
            TabIndex        =   244
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
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
            Format          =   135790593
            CurrentDate     =   42877
         End
         Begin VB.TextBox txtProrroga 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   4320
            MaxLength       =   80
            TabIndex        =   208
            Top             =   1940
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox chk_pago_mensual 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "PAGO MENSUAL"
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
            Height          =   220
            Left            =   1440
            TabIndex        =   168
            Top             =   2445
            Width           =   1455
         End
         Begin VB.CheckBox chk_anual 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "PAGO ANUAL"
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
            Height          =   220
            Left            =   1440
            TabIndex        =   167
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txt_precio_plan 
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
            Left            =   1440
            MaxLength       =   80
            TabIndex        =   166
            Top             =   640
            Width           =   1095
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesarplan 
            Height          =   460
            Left            =   6120
            TabIndex        =   169
            Top             =   2470
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   820
            BTYPE           =   5
            TX              =   "GUARDAR PLAN"
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
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":9B50
            PICN            =   "FrmDetalleCliente.frx":9B6C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcSucursal 
            Height          =   315
            Left            =   1440
            TabIndex        =   171
            Top             =   1005
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin VitekeySoft.ChameleonBtn cmdcerrardetalle 
            Height          =   250
            Left            =   8140
            TabIndex        =   181
            Top             =   120
            Width           =   200
            _ExtentX        =   344
            _ExtentY        =   450
            BTYPE           =   5
            TX              =   "X"
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
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleCliente.frx":D1B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcPlanServicio 
            Height          =   315
            Left            =   1440
            TabIndex        =   243
            Top             =   285
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin MSComCtl2.DTPicker DtpFechaCorte 
            Height          =   300
            Left            =   6960
            TabIndex        =   269
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
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
            Format          =   135790593
            CurrentDate     =   42877
         End
         Begin MSDataListLib.DataCombo DtcMedioContacto 
            Height          =   315
            Left            =   6120
            TabIndex        =   274
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin MSDataListLib.DataCombo DtcEstadoEmpresa 
            Height          =   315
            Left            =   6120
            TabIndex        =   275
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin MSDataListLib.DataCombo DtcPrioridad 
            Height          =   315
            Left            =   6120
            TabIndex        =   280
            Top             =   2040
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
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
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRIORIDAD :"
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
            Index           =   7
            Left            =   5280
            TabIndex        =   282
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "FECHA SUSCRRIPCION :"
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
            Index           =   4
            Left            =   1380
            TabIndex        =   281
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ESTADO :"
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
            Index           =   6
            Left            =   5520
            TabIndex        =   277
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MEDIO :"
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
            Index           =   5
            Left            =   5580
            TabIndex        =   276
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "FECHA CORTE:"
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
            Height          =   225
            Index           =   3
            Left            =   5400
            TabIndex        =   268
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "ULT.DIA PAGO"
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
            Height          =   225
            Index           =   1
            Left            =   3000
            TabIndex        =   264
            Top             =   2205
            Width           =   1215
         End
         Begin VB.Label lblprorroga 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "PRORROGA :"
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
            Height          =   225
            Left            =   3000
            TabIndex        =   209
            Top             =   1965
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUCURSAL :"
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
            Index           =   0
            Left            =   465
            TabIndex        =   170
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label LblObservacion 
            Alignment       =   2  'Center
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
            Index           =   22
            Left            =   30
            TabIndex        =   165
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label LblObservacion 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLAN SERVICIO :"
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
            Index           =   21
            Left            =   240
            TabIndex        =   164
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtLicencia 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   154
         Top             =   2460
         Width           =   1695
      End
      Begin VB.TextBox txtpeso 
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
         Left            =   8835
         MaxLength       =   80
         TabIndex        =   136
         Text            =   "0.1"
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txttalla 
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
         Left            =   8835
         MaxLength       =   80
         TabIndex        =   135
         Text            =   "0.1"
         Top             =   1740
         Width           =   615
      End
      Begin VB.TextBox txtalergias 
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
         Left            =   6840
         MaxLength       =   80
         TabIndex        =   134
         Top             =   2050
         Width           =   2655
      End
      Begin VB.TextBox txtvacunas 
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
         Left            =   6840
         MaxLength       =   80
         TabIndex        =   131
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox txtenfermedades 
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
         Left            =   6825
         MaxLength       =   80
         TabIndex        =   129
         Top             =   1400
         Width           =   1935
      End
      Begin VB.CheckBox chk_mama 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "MAM�"
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
         Left            =   -70080
         TabIndex        =   127
         Top             =   3060
         Width           =   855
      End
      Begin VB.CheckBox chk_papa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PAP�"
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
         Left            =   -71040
         TabIndex        =   126
         Top             =   3060
         Width           =   855
      End
      Begin VB.CheckBox chkseguro 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SEGURO [OPCIONAL]"
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
         Height          =   220
         Left            =   2160
         TabIndex        =   124
         Top             =   2580
         Width           =   3255
      End
      Begin VB.CheckBox chkrecuperacion_no 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NO"
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
         Height          =   200
         Left            =   2880
         TabIndex        =   122
         Top             =   1500
         Width           =   570
      End
      Begin VB.CheckBox chkrecuperacion_si 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SI"
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
         Height          =   200
         Left            =   2160
         TabIndex        =   121
         Top             =   1500
         Width           =   570
      End
      Begin VB.TextBox txthabilidad 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   118
         Top             =   2220
         Width           =   3255
      End
      Begin VB.TextBox txttercio 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   117
         Top             =   1845
         Width           =   3255
      End
      Begin VB.TextBox txtpromovido 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   116
         Top             =   1140
         Width           =   3255
      End
      Begin VB.TextBox txtieprocedencia 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   115
         Top             =   780
         Width           =   3255
      End
      Begin VB.TextBox txtdireccion_familia 
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
         Left            =   -71055
         TabIndex        =   104
         Top             =   2655
         Width           =   2295
      End
      Begin VB.CheckBox chkclientemayor 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLIENTE X MAYOR"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   86
         Top             =   2460
         Width           =   2895
      End
      Begin VB.TextBox txtDni 
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
         Left            =   -73815
         TabIndex        =   70
         Top             =   1980
         Width           =   1575
      End
      Begin VB.TextBox TxtFmaterno 
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
         Left            =   -73815
         TabIndex        =   69
         Top             =   2580
         Width           =   1575
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "AGREGAR ITEM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   -67455
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2715
         Width           =   1935
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
         Left            =   -71055
         TabIndex        =   67
         Top             =   2355
         Width           =   2295
      End
      Begin VB.TextBox TxtFpaterno 
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
         Left            =   -73815
         TabIndex        =   66
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtFnombers 
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
         Left            =   -73815
         TabIndex        =   65
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdeliminarcuenta 
         Caption         =   "QUITAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69840
         TabIndex        =   59
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdagregarcuenta 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71040
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtnumerocuenta 
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
         Left            =   -71040
         MaxLength       =   50
         TabIndex        =   57
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox TxtFono 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -68280
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ELIMINAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68280
         TabIndex        =   30
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66840
         TabIndex        =   29
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txtMaximoCredito 
         Alignment       =   1  'Right Justify
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
         Left            =   -72600
         MaxLength       =   11
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtRucEmpresa 
         Alignment       =   1  'Right Justify
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
         Left            =   -72600
         MaxLength       =   11
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OptSincredito 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SIN LINEA DE CREDITO"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   2160
         Width           =   2895
      End
      Begin VB.OptionButton ChkMaximoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASIGNAR CREDITO     :"
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   1560
         Width           =   2175
      End
      Begin VB.OptionButton chkEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMPRESA VINCULADA :"
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTelefonos 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   8388608
         FocusRect       =   0
         GridLines       =   2
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
      Begin MSDataListLib.DataCombo DtcArea 
         Height          =   330
         Left            =   -68280
         TabIndex        =   44
         Top             =   780
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSDataListLib.DataCombo DtcBanco 
         Height          =   330
         Left            =   -74070
         TabIndex        =   60
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   330
         Left            =   -74070
         TabIndex        =   61
         Top             =   1065
         Width           =   2655
         _ExtentX        =   4683
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFamiliares 
         Height          =   1410
         Left            =   -74880
         TabIndex        =   64
         Top             =   420
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2487
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
      Begin MSDataListLib.DataCombo DtcParentesco 
         Height          =   330
         Left            =   -71055
         TabIndex        =   71
         Top             =   2010
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuentas 
         Height          =   1335
         Left            =   -74040
         TabIndex        =   102
         Top             =   1500
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   8388608
         FocusRect       =   0
         GridLines       =   2
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
      Begin MSDataListLib.DataCombo DtcOcupacion 
         Height          =   330
         Left            =   -67440
         TabIndex        =   107
         Top             =   1890
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataListLib.DataCombo DtcGradoinstruccion 
         Height          =   330
         Left            =   -67440
         TabIndex        =   108
         Top             =   2340
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSDataListLib.DataCombo DtcNivel 
         Height          =   315
         Left            =   6825
         TabIndex        =   120
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcTiponacimiento 
         Height          =   315
         Left            =   6825
         TabIndex        =   139
         Top             =   1050
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcGrado 
         Height          =   315
         Left            =   6825
         TabIndex        =   148
         Top             =   705
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   2175
         TabIndex        =   149
         Top             =   420
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSDataListLib.DataCombo DtcMatricula 
         Height          =   315
         Left            =   6495
         TabIndex        =   151
         Top             =   2370
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfplanservicio 
         Height          =   2490
         Left            =   -74880
         TabIndex        =   159
         Top             =   660
         Width           =   8360
         _ExtentX        =   14737
         _ExtentY        =   4392
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
      Begin VitekeySoft.ChameleonBtn cmdnuevoplan 
         Height          =   405
         Left            =   -66440
         TabIndex        =   160
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "NUEVO PLAN"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":D1D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdmodificarplan 
         Height          =   405
         Left            =   -66440
         TabIndex        =   161
         Top             =   1140
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "MODIFICAR "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":D1EC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdeliminarplan 
         Height          =   405
         Left            =   -66440
         TabIndex        =   162
         Top             =   1620
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "ELIMINAR"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":D208
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo Dtcseguroempresa 
         Height          =   330
         Left            =   -70680
         TabIndex        =   173
         Top             =   1020
         Width           =   5055
         _ExtentX        =   8916
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
      Begin MSComCtl2.DTPicker DtpFechaCaducidad_transporte 
         Height          =   345
         Left            =   -69480
         TabIndex        =   179
         Top             =   2220
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   135790593
         CurrentDate     =   42868
      End
      Begin VitekeySoft.ChameleonBtn cmdActualizar_password 
         Height          =   315
         Left            =   -66840
         TabIndex        =   214
         Top             =   1440
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "ACTUALIZAR"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleCliente.frx":D224
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcPension 
         Height          =   315
         Left            =   6495
         TabIndex        =   227
         Top             =   2700
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcEstado 
         Height          =   315
         Left            =   6495
         TabIndex        =   233
         Top             =   3030
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   8438015
         ForeColor       =   8388608
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
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIONES :"
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
         Index           =   0
         Left            =   -69840
         TabIndex        =   279
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO :"
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
         Index           =   20
         Left            =   5790
         TabIndex        =   232
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENSION :"
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
         Index           =   19
         Left            =   5775
         TabIndex        =   228
         Top             =   2745
         Width           =   705
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
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
         Index           =   2
         Left            =   -69840
         TabIndex        =   212
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblidtelefono 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69600
         TabIndex        =   184
         Top             =   2340
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.CADUCIDAD :"
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
         Left            =   -70680
         TabIndex        =   176
         Top             =   2220
         Width           =   1035
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.EMISION :"
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
         Left            =   -70440
         TabIndex        =   175
         Top             =   1860
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� POLIZA :"
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
         Left            =   -70410
         TabIndex        =   174
         Top             =   1500
         Width           =   765
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LICENCIA CONDUCIR :"
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
         Index           =   1
         Left            =   -74040
         TabIndex        =   155
         Top             =   2580
         Width           =   1485
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MATRICULA :"
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
         Index           =   18
         Left            =   5535
         TabIndex        =   152
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO :"
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
         Index           =   17
         Left            =   1320
         TabIndex        =   150
         Top             =   480
         Width           =   705
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESO "
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
         Index           =   14
         Left            =   8985
         TabIndex        =   138
         Top             =   840
         Width           =   405
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TALLA:"
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
         Index           =   13
         Left            =   8910
         TabIndex        =   137
         Top             =   1455
         Width           =   495
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALERGIAS :"
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
         Index           =   12
         Left            =   5895
         TabIndex        =   133
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VACUNAS :"
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
         Index           =   11
         Left            =   5910
         TabIndex        =   132
         Top             =   1815
         Width           =   795
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENFERMEDAD :"
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
         Index           =   10
         Left            =   5640
         TabIndex        =   130
         Top             =   1455
         Width           =   1065
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO NACI :"
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
         Index           =   9
         Left            =   5910
         TabIndex        =   128
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIVE CON :"
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
         Left            =   -71835
         TabIndex        =   125
         Top             =   3060
         Width           =   765
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRADO :"
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
         Index           =   8
         Left            =   6090
         TabIndex        =   123
         Top             =   800
         Width           =   615
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL :"
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
         Index           =   5
         Left            =   6210
         TabIndex        =   119
         Top             =   420
         Width           =   495
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HABILIDAD :"
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
         Index           =   7
         Left            =   1140
         TabIndex        =   114
         Top             =   2220
         Width           =   885
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERCIO ESTUDIANTIL :"
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
         Index           =   6
         Left            =   540
         TabIndex        =   113
         Top             =   1860
         Width           =   1485
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REQUIERE RECUPERACION :"
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
         Index           =   4
         Left            =   150
         TabIndex        =   112
         Top             =   1515
         Width           =   1875
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROMOVIDO :"
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
         Index           =   3
         Left            =   1020
         TabIndex        =   111
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.E PROCEDENCIA :"
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
         Index           =   2
         Left            =   750
         TabIndex        =   110
         Top             =   855
         Width           =   1275
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.INSTRUCC :"
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
         Left            =   -68505
         TabIndex        =   106
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OCUPACION :"
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
         Left            =   -68445
         TabIndex        =   105
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION :"
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
         Left            =   -71925
         TabIndex        =   103
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI :"
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
         Index           =   23
         Left            =   -74340
         TabIndex        =   77
         Top             =   1980
         Width           =   345
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENTESCO :"
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
         Left            =   -72045
         TabIndex        =   76
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO :"
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
         Left            =   -71865
         TabIndex        =   75
         Top             =   2355
         Width           =   795
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.MATERNO :"
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
         Index           =   25
         Left            =   -74940
         TabIndex        =   74
         Top             =   2655
         Width           =   945
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.PATERNO :"
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
         Index           =   24
         Left            =   -74880
         TabIndex        =   73
         Top             =   2340
         Width           =   885
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES :"
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
         Index           =   26
         Left            =   -74790
         TabIndex        =   72
         Top             =   2925
         Width           =   795
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         Left            =   -74805
         TabIndex        =   63
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO :"
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
         Left            =   -74775
         TabIndex        =   62
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AREA :"
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
         Left            =   -69510
         TabIndex        =   45
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� TELEFONICO :"
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
         Left            =   -69600
         TabIndex        =   43
         Top             =   1500
         Width           =   1125
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   1815
         Left            =   -69720
         Top             =   480
         Width           =   4215
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   -70920
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label lblEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Height          =   225
         Left            =   -74760
         TabIndex        =   27
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   -74880
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   -72105
         TabIndex        =   21
         Top             =   2160
         Width           =   225
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   2175
         Index           =   1
         Left            =   -70080
         Top             =   480
         Width           =   4575
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000005&
         Height          =   1020
         Left            =   5520
         Top             =   2360
         Width           =   4095
      End
   End
   Begin MSDataListLib.DataCombo DtcDistrito 
      Height          =   330
      Left            =   1320
      TabIndex        =   32
      Top             =   8865
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
      Left            =   1320
      TabIndex        =   33
      Top             =   8160
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
      Left            =   1320
      TabIndex        =   34
      Top             =   8520
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
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   6240
      TabIndex        =   35
      Top             =   8865
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
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
   Begin VitekeySoft.TextBoxPlus txtRazonSocial 
      Height          =   315
      Left            =   1560
      TabIndex        =   78
      Top             =   1660
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      BackColor       =   12648447
      BackColorEnabled=   12648447
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
      ForeColor       =   4194304
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDireccion 
      Height          =   810
      Left            =   1560
      TabIndex        =   87
      Top             =   2400
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1429
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
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
   Begin VitekeySoft.ChameleonBtn cmdmodificar_direccion 
      Height          =   300
      Left            =   7155
      TabIndex        =   89
      Top             =   2970
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      BTYPE           =   5
      TX              =   "MODIFICA"
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
      BCOL            =   8438015
      BCOLO           =   8438015
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleCliente.frx":D240
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar_direccion 
      Height          =   300
      Left            =   7155
      TabIndex        =   90
      Top             =   3285
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      BCOL            =   8438015
      BCOLO           =   8438015
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleCliente.frx":D25C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPais 
      Height          =   330
      Left            =   1320
      TabIndex        =   158
      Top             =   7815
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
   Begin MSDataListLib.DataCombo DtcCobertura 
      Height          =   330
      Left            =   6240
      TabIndex        =   242
      Top             =   7815
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DtcZona 
      Height          =   330
      Left            =   6240
      TabIndex        =   247
      Top             =   8520
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   330
      Left            =   6645
      TabIndex        =   249
      Top             =   7440
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   700
      Left            =   9600
      TabIndex        =   250
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1244
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
      MICON           =   "FrmDetalleCliente.frx":D278
      PICN            =   "FrmDetalleCliente.frx":D294
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   700
      Left            =   10515
      TabIndex        =   251
      Top             =   8520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1244
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
      MICON           =   "FrmDetalleCliente.frx":108DC
      PICN            =   "FrmDetalleCliente.frx":108F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTipoZona 
      Height          =   330
      Left            =   6240
      TabIndex        =   253
      Top             =   8175
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DtcFrecuenciaVisita 
      Height          =   315
      Left            =   9600
      TabIndex        =   256
      Top             =   8160
      Width           =   1720
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
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
   Begin MSDataListLib.DataCombo DtcGiro 
      Height          =   315
      Left            =   9600
      TabIndex        =   258
      Top             =   7820
      Width           =   1720
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
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
   Begin MSDataListLib.DataCombo DtcSexo 
      Height          =   330
      Left            =   4920
      TabIndex        =   289
      Top             =   3240
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label lblCodigoUbigeoSunat 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   287
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   4
      Left            =   5340
      TabIndex        =   262
      Top             =   8880
      Width           =   855
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " ZONA :"
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
      Index           =   0
      Left            =   5700
      TabIndex        =   261
      Top             =   8595
      Width           =   495
   End
   Begin VB.Label lblprovincia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GIRO :"
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
      Index           =   2
      Left            =   9150
      TabIndex        =   260
      Top             =   7920
      Width           =   435
   End
   Begin VB.Label lblDepartamento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTAMENTO :"
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
      Index           =   1
      Left            =   60
      TabIndex        =   259
      Top             =   8200
      Width           =   1215
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.VISITA :"
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
      Index           =   3
      Left            =   8940
      TabIndex        =   257
      Top             =   8265
      Width           =   645
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA PRECIO :"
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
      Index           =   2
      Left            =   5220
      TabIndex        =   254
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO ZONA :"
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
      Index           =   1
      Left            =   5370
      TabIndex        =   252
      Top             =   8280
      Width           =   825
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FRE.VISITAS:"
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
      Left            =   5565
      TabIndex        =   248
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label lbl_id_matricula 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   9840
      TabIndex        =   231
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbldistrito 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISTRITO :"
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
      TabIndex        =   100
      Top             =   8930
      Width           =   705
   End
   Begin VB.Label lblprovincia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA :"
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
      Index           =   0
      Left            =   405
      TabIndex        =   99
      Top             =   8600
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.NACIMIENTO :"
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
      Left            =   360
      TabIndex        =   85
      Top             =   3240
      Width           =   1125
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A. MATERNO :"
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
      Left            =   510
      TabIndex        =   39
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A. PATERNO :"
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
      Left            =   570
      TabIndex        =   38
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRES :"
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
      Left            =   690
      TabIndex        =   37
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label LblCodPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   120
      TabIndex        =   19
      Top             =   75
      Width           =   1935
   End
   Begin VB.Label LblTipoDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI:"
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
      Left            =   2355
      TabIndex        =   14
      Top             =   195
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   8280
      Picture         =   "FrmDetalleCliente.frx":10C12
      Stretch         =   -1  'True
      Top             =   200
      Width           =   2415
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL :"
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
      Left            =   900
      TabIndex        =   10
      Top             =   3660
      Width           =   585
   End
   Begin VB.Label LblTelefono1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel�fono 1 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -2910
      TabIndex        =   9
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel�fono 2 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -1560
      TabIndex        =   8
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label LblFax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -1200
      TabIndex        =   7
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Left            =   630
      TabIndex        =   3
      Top             =   2100
      Width           =   855
   End
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE COMPLETO :"
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
      Left            =   -30
      TabIndex        =   2
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2820
      Left            =   8160
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmDetallePersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Dim TipoDocumento As String
Dim strCodPersona As String
Dim StrCliente As String, strProveedor As String, Per_N As String, StrPercepcion As String, StrRetencion As String, StrAuspiciador As String
Dim StrContable As String, StrTransporte As String, StrPersonal As String, StrAlmacen As String
Public img As String
Dim descuento_por As Single
Dim Adelantado As Double
Public IdEmpresa As Long
Public RucExt As String
Public Procedencia As EnumProcede
Private WithEvents cSunat As cls_QrySUNAT
Attribute cSunat.VB_VarHelpID = -1



Private Function validacion_extrema(ByVal in_dni As String) As Boolean
Dim in_mensaje As String
Dim j As Integer
in_mensaje = ""
j = 0
If Trim(Me.txtRuc.Text) = "" Then
   j = 1
   in_mensaje = "[ " & j & " ] Debe Ingresar un DNI/RUC Valido"
End If

If Trim(Me.TxtRazonSocial.Text) = "" Then
    If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Debe Ingresar una Nombre/Razon Social"
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Debe Ingresar una Nombre/Razon Social"
    End If
End If


If Trim(Me.txtEmail.Text) = "" Then
    If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Debe de Ingresar un E-mail."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Debe de Ingresar un E-mail."
    End If
    
End If

If Trim(Me.TxtDireccion1.Text) = "" Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Ingrese una Direccion Valida."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Ingrese una Direccion Valida."
    End If
    
    
End If

If get_telefono_last(in_dni) = "" Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Debe de Ingresar un Numero Telef�nico."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Debe de Ingresar un Numero Telef�nico."
    End If
    
End If


If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), "@") < 1 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Ingrese un mail Valido."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Ingrese un mail Valido."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), ",") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), ";") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), "-") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If

If in_mensaje <> "" Then
    MsgBox in_mensaje, vbInformation, KEY_VENDEDOR
    validacion_extrema = False
Else
    validacion_extrema = True
End If

End Function


Private Function validacion_simple(ByVal in_dni As String) As Boolean
Dim in_mensaje As String
Dim j As Integer
in_mensaje = ""
j = 0
If Trim(Me.txtRuc.Text) = "" Then
   j = 1
   in_mensaje = "[ " & j & " ] Debe Ingresar un DNI/RUC Valido"
End If

If Trim(Me.TxtRazonSocial.Text) = "" Then
    If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Debe Ingresar una Nombre/Razon Social"
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Debe Ingresar una Nombre/Razon Social"
    End If
End If



If Trim(Me.TxtDireccion1.Text) = "" Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Ingrese una Direccion Valida."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Ingrese una Direccion Valida."
    End If
    
    
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), "@") < 1 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Ingrese un mail Valido."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Ingrese un mail Valido."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), ",") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), ";") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If

If Len(Me.txtEmail.Text) > 0 Then
If InStr(1, Trim(Me.txtEmail.Text), "-") > 0 Then
   If j = 0 Then
        j = 1
        in_mensaje = "[ " & j & " ] Solo se Permite un Mail por cliente."
    Else
        j = j + 1
        in_mensaje = in_mensaje & Chr(13) & "[ " & j & " ] Solo se Permite un Mail por cliente."
    End If
End If
End If
If in_mensaje <> "" Then
    MsgBox in_mensaje, vbInformation, KEY_VENDEDOR
    validacion_simple = False
Else
    validacion_simple = True
End If

End Function


Private Sub Save()
Dim StrNombre As String
Dim StrDireccion As String
Dim fecha_nacimiento As String
Adelantado = 0
StrNombre = Replace(Trim(Me.TxtRazonSocial.Text), "'", " ")   'Comillas(Trim(Me.txtRazonSocial.Text))
StrDireccion = Replace(Trim(Me.TxtDireccion1.Text), "'", " ")
  If Me.chk_igv.Value = 1 Then
    in_afecto_igv = "si"
  Else
    in_afecto_igv = "no"
  
  End If
  
  
  
  
  If KEY_VALIDACION_EXTREMA = "si" Then
    If validacion_extrema(Me.txtRuc.Text) = False Then
        Exit Sub
     End If
  Else
      If validacion_simple(Me.txtRuc.Text) = False Then
        Exit Sub
     End If
  End If
  
  
      
      
      Call verificaTipo
      If Me.chk_extranjeria.Value = 1 Then
         in_extranjero = "si"
      Else
         in_extranjero = "no"
      End If
      
     
      
      strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
      Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & Trim(Me.txtRuc.Text) & "' " & _
                ",'" & Replace(UCase(Me.txtPaterno.Text), "'", " ") & "', " & _
                "'" & Replace(UCase(Me.txtMaterno.Text), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(Me.txtNombre.Text)), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(Me.TxtRazonSocial.Text)), "'", " ") & "' " & _
                ",'" & Trim(Me.TxtDireccion1.Text) & "' " & _
                ",'" & Trim(Me.txtTelefono.Text) & "' " & _
                ",'" & Me.txtEmail.Text & "'" & _
                ",'" & StrTransporte & "' " & _
                ",'" & StrContable & "'" & _
                ",'" & strProveedor & "' " & _
                ",'" & StrPersonal & "' " & _
                ",'" & StrAuspiciador & "' " & _
                ",'" & StrAlmacen & "' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                 
                
                strCadena = "UPDATE persona SET sexo='" & Me.DtcSexo.BoundText & "', codigo_ubigeo_sunat='" & Me.lblCodigoUbigeoSunat.Caption & "',afecto_percepcion='" & StrPercepcion & "',extranjero='" & in_extranjero & "',id_pais='" & Me.DtcPais.BoundText & "', peso='" & Val(Me.txtpeso.Text) & "',estatura='" & Trim(Me.txttalla.Text) & "',a_paterno='" & Trim(Me.txtPaterno.Text) & "',a_materno='" & Trim(Me.txtMaterno.Text) & "',nombres='" & Trim(Me.txtNombre.Text) & "',id_dia='" & Trim(Me.txtdia.Text) & "',id_mes='" & Trim(Me.DtcMes.BoundText) & "',id_anio='" & Trim(Me.txtAnio.Text) & "'," & _
                "nombre_completo='" & Replace(Trim(Me.TxtRazonSocial.Text), "'", " ") & "',celular='" & get_telefono_last(Trim(Me.txtRuc.Text)) & "',direccion='" & Replace(Trim(Me.TxtDireccion1.Text), "'", " ") & "',licencia='" & Trim(Me.TxtLicencia.Text) & "',mail='" & Trim(Me.txtEmail.Text) & "',id_departamento='" & Me.DtcDepartamento.BoundText & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_distrito='" & Me.DtcDistrito.BoundText & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                CnBd.Execute (strCadena)
                
                If StrAlmacen = "si" Then
                    Call persona_almacen(Trim(Me.txtRuc.Text))
                End If
                If img <> vbNullString And Trim(str(Me.txtRuc.Text)) <> vbNullString Then
                    strCadena = "UPDATE persona SET foto='" & img & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                    CnBd.Execute (strCadena)
                    img = ""
                End If
                Else
                    cPersona = Trim(Me.LblCodPersona.Caption)
                    If get_cliente_uso(cPersona) = True Then
                       strCadena = "UPDATE persona SET sexo='" & Me.DtcSexo.BoundText & "',codigo_ubigeo_sunat='" & Me.lblCodigoUbigeoSunat.Caption & "',afecto_percepcion='" & StrPercepcion & "',extranjero='" & in_extranjero & "',celular='" & get_telefono_last(Trim(Me.txtRuc.Text)) & "',id_pais='" & Me.DtcPais.BoundText & "',peso='" & Val(Me.txtpeso.Text) & "',estatura='" & Trim(Me.txttalla.Text) & "',a_paterno='" & Trim(Me.txtPaterno.Text) & "',a_materno='" & Trim(Me.txtMaterno.Text) & "',nombres='" & Trim(Me.txtNombre.Text) & "',id_dia='" & Trim(Me.txtdia.Text) & "',id_mes='" & Trim(Me.DtcMes.BoundText) & "',id_anio='" & Trim(Me.txtAnio.Text) & "'," & _
                       "direccion='" & Trim(Me.TxtDireccion1.Text) & "',licencia='" & Trim(Me.TxtLicencia.Text) & "',mail='" & Trim(Me.txtEmail.Text) & "',id_departamento='" & Me.DtcDepartamento.BoundText & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_distrito='" & Me.DtcDistrito.BoundText & "' WHERE dni='" & cPersona & "'"
                    Else
                       strCadena = "UPDATE persona SET sexo='" & Me.DtcSexo.BoundText & "',codigo_ubigeo_sunat='" & Me.lblCodigoUbigeoSunat.Caption & "',afecto_percepcion='" & StrPercepcion & "',extranjero='" & in_extranjero & "',celular='" & get_telefono_last(Trim(Me.txtRuc.Text)) & "',id_pais='" & Me.DtcPais.BoundText & "',peso='" & Val(Me.txtpeso.Text) & "',estatura='" & Trim(Me.txttalla.Text) & "',a_paterno='" & Trim(Me.txtPaterno.Text) & "',a_materno='" & Trim(Me.txtMaterno.Text) & "',nombres='" & Trim(Me.txtNombre.Text) & "',id_dia='" & Trim(Me.txtdia.Text) & "',id_mes='" & Trim(Me.DtcMes.BoundText) & "',id_anio='" & Trim(Me.txtAnio.Text) & "'," & _
                       "nombre_completo='" & Trim(Me.TxtRazonSocial.Text) & "',direccion='" & Trim(Me.TxtDireccion1.Text) & "',licencia='" & Trim(Me.TxtLicencia.Text) & "',mail='" & Trim(Me.txtEmail.Text) & "',id_departamento='" & Me.DtcDepartamento.BoundText & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_distrito='" & Me.DtcDistrito.BoundText & "' WHERE dni='" & cPersona & "'"
                    End If
                    
                    
                    CnBd.Execute (strCadena)
                
                    
                    
                    strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & cPersona & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstT(strCadena)
                    If rstT.RecordCount < 1 Then
                        strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_almacen,passwordaccesso)VALUES ('" & cPersona & "','" & KEY_RUC & "','" & StrAlmacen & "','" & cPersona & "')"
                        CnBd.Execute (strCadena)
                    End If
                  
                End If
           ' If img <> vbNullString And Trim(Str(Me.txtRUC.text)) <> vbNullString Then
                'Ret = Guardar_Imagen(CnBd, "SELECT per_foto From Persona Where cPersona=" & Trim(cPersona), "per_foto", img)
             
            'End If
             
                    If Me.ChkMaximoCredito.Value = True Then
                        strcredito = "si"
                    Else
                        strcredito = "no"
                    End If
                    
                    If Me.chkclientemayor.Value = 1 Then
                        strmayor = "si"
                    Else
                        strmayor = "no"
                    End If
                    
                    If Me.chkhabilitado.Value = 1 Then
                        in_habilitado = "si"
                    Else
                        in_habilitado = "no"
                    End If
                    If StrCliente = "no" Then
                        If MsgBox("Esta Seguro de NO registrar como CLIENTE", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbNo Then
                            Exit Sub
                        End If
                        
                    End If
                    
                    strCadena = "UPDATE entidad_empresa SET observacion='" & Replace(Me.txtObservacion.Text, "'", "") & "',id_zona='" & Val(Me.DtcZona.BoundText) & "',id_percepcion='" & StrPercepcion & "', id_tipo_cliente='" & Me.DtcCobertura.BoundText & "', afecto_igv='" & in_afecto_igv & "',id_cliente='" & StrCliente & "', habilitado='" & in_habilitado & "', id_cliente='" & StrCliente & "', cliente_mayor='" & strmayor & "', id_credito='" & strcredito & "',monto_credito='" & Val(Me.txtMaximoCredito.Text) & "',id_transporte='" & StrTransporte & "',id_proveedor='" & strProveedor & "',id_vendedor='" & Me.DtcVendedor.BoundText & "',id_almacen='" & StrAlmacen & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                
                
                
                If Me.ChkTransporte.Value = 1 Then
                    strCadena = "UPDATE persona SET licencia='" & Me.TxtLicencia.Text & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                    CnBd.Execute (strCadena)
                 
                End If
             
                
               ' If Me.ChkPersonal.Value = 1 Then
                '     strCadena = "UPDATE entidad_empresa SET id_personal='" & StrPersonal & "',password='" & Trim(Me.TxtPassword.Text) & "',passwordaccesso='" & Trim(Me.TxtPassword.Text) & "',id_planilla='" & Me.DtcPlanilla.Text & "',id_sucursal='" & Me.DtcSucursal.BoundText & "',sueldo='" & Val(Me.txtSueldoMensual.Text) & "',id_proveedor='" & strProveedor & "',id_afp='" & Me.DtcAfp.BoundText & "',rta_quinta='" & Val(Me.TxtRentaquinta.Text) & "',asig_familiar='" & Val(Me.txtAsiganacion_familiar.Text) & "',bonificacion_extraordinaria='" & Val(Me.TxtBonificacion.Text) & "',cuspp='" & Val(Me.TxtCuspp.Text) & "',essalud='" & Val(Me.TxtEssalud.Text) & "',sndp='" & Val(Me.TxtSNDP.Text) & "',fecha_ingreso='" & Format(Me.DtpIngreso.Value, "YYYY-mm-dd") & "',id_cargo='" & Me.DtcCargo.BoundText & "',id_condicion='" & Me.DtcRegimen.BoundText & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                 '    CnBd.Execute (strCadena)
                'Else
                  '  strCadena = "UPDATE entidad_empresa SET id_personal='" & StrPersonal & "',id_planilla='" & DtcPlanilla.Text & "',id_sucursal='" & Me.DtcSucursal.BoundText & "',sueldo='" & Val(Me.txtSueldoMensual.Text) & "',id_proveedor='" & strProveedor & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                 '   CnBd.Execute (strCadena)
                'End If
                    
                
                
               ' If Me.DtcDistrito.BoundText <> "" Then
                '    strCadena = "UPDATE persona SET id_distrito='" & Me.DtcDistrito.BoundText & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_departamento='" & Me.DtcDepartamento.BoundText & "', id_urbanizacion='" & Me.DtcUrbanizacion.BoundText & "',id_zona='" & Me.DtcZona.BoundText & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                 '   CnBd.Execute (strCadena)
                'End If
                
                
    
           If Me.chk_seguro_transporte.Value = 1 Then
              If Me.chk_estadoseguro.Value = 1 Then
                 in_estado_seguro = "si"
              Else
                in_estado_seguro = "no"
              End If
              
              strCadena = "CALL p_put_seguro_persona('" & Trim(Me.txtRuc.Text) & "','" & Me.Dtcseguroempresa.BoundText & "','" & Trim(Me.txtpolizaseguro_transporte.Text) & "','" & Format(Me.DtpFechaEmision_transporte.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaCaducidad_transporte.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & in_estado_seguro & "','" & Me.Dtcseguroempresa.Text & "','" & KEY_RUC & "')"
              CnBd.Execute (strCadena)
           End If
           
           
           If KEY_RUBRO = "00025" Then
            Call put_estudiante(Trim(Me.txtRuc.Text))
            strCadena = "SELECT * FROM college_matricula WHERE dni='" & Trim(Me.txtRuc.Text) & "'  and id_periodo='" & Me.DtcPeriodo.BoundText & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                in_matricula = rstL("id_matricula")
            Else
                in_matricula = 0
            End If
            If Me.chk_beca.Value = 1 Then
               in_beca = "si"
            Else
               in_beca = "no"
            End If
            If Me.chk_media_beca.Value = 1 Then
               in_media_beca = "si"
            Else
                in_media_beca = "no"
            End If
                
                strCadena = "call put_matricula_college('" & in_matricula & "','" & Trim(Me.txtRuc.Text) & "','" & Me.DtcGrado.BoundText & "','" & in_beca & "','" & in_media_beca & "','" & Me.DtcMatricula.BoundText & "','" & Me.DtcPension.BoundText & "','" & Me.DtcEstado.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            
            End If
                
        
        
        
        
        
        If KEY_RUBRO = "00027" Then
               Call put_biblio(Trim(Me.txtRuc.Text), Me.DtcFacultad.BoundText, Me.DtcTipoAcceso.BoundText, Me.DtcCiclo.BoundText)
        End If
                
                
               
    
       
     If FrmPersona.Procedencia = modificar Then
        FrmPersona.Procedencia = Neutro
        If KEY_RUBRO = "00025" Then
            strCadena = "SELECT * FROM view_estudiante WHERE dni = '" & Trim(Me.txtRuc.Text) & "' and  id_empresa='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT * FROM view_cliente WHERE id_cliente='si' and dni = '" & Trim(Me.txtRuc.Text) & "' and  id_empresa='" & KEY_RUC & "'"
        End If
        
        
        
        Call FrmPersona.llenarGrid(FrmPersona.HfdPersona)
        Unload Me
        Exit Sub
     End If
     
     If FrmMatricula.Procedencia = modificar Then
        FrmMatricula.Procedencia = Neutro
        strCadena = "SELECT * FROM view_matricula WHERE dni = '" & Trim(Me.txtRuc.Text) & "' and  ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'"
        Call FrmMatricula.llenarGrid(FrmMatricula.HfdPersona)
        Unload Me
        Exit Sub
     End If
     
     If FrmPersona.Procedencia = nuevo Then
        Call FrmPersona.actualizar
        FrmPersona.Procedencia = Neutro
        Call Resalta(FrmPersona.txtRuc)
        Unload Me
        Exit Sub
     End If
       
       
       If FrmVentas.Procedencia = nuevo Or FrmVentas.Procedencia = Selecionar Then
            FrmVentas.Procedencia = Neutro
            If Len(Trim(Me.txtRuc.Text)) = 11 Then
                FrmVentas.TxtCodCliente.Text = Trim(Me.txtRuc.Text)
            End If
            FrmVentas.TxtCliente.Text = Trim(Me.TxtRazonSocial.Text)
            FrmVentas.txtDireccion.Text = Trim(Me.TxtDireccion1.Text)
            FrmVentas.precionar_cliente
            Call Resalta(FrmVentas.TxtCodProducto)
            Unload Me
            Exit Sub
        End If
       If FrmComprasGastos.Procedencia = nuevo Then
          FrmComprasGastos.txtDni.Text = Me.txtRuc.Text
          FrmComprasGastos.lblcliente.Caption = UCase(Me.TxtRazonSocial.Text)
          FrmComprasGastos.Procedencia = Neutro
          Unload Me
          Exit Sub
       End If
        
        
        
        If FrmSolicitudViaticosDeclarar.Procedencia = nuevo Then
            If Len(Trim(Me.txtRuc.Text)) = 11 Then
                FrmSolicitudViaticosDeclarar.txtRuc.Text = Me.txtRuc.Text
                'FrmSolicitudViaticosDeclarar.lblRazonSocial.Caption = Trim(Me.TxtRazonsocial.text)
                FrmSolicitudViaticosDeclarar.Procedencia = Neutro
                
            End If
            
            Unload Me
            Exit Sub
        End If
        If FrmPersona.Procedencia = nuevo Then
            
            Call FrmPersona.actualizar
            Unload Me
            Exit Sub
        End If
       
        If FrmCompras.Procedencia = nuevo Then
            FrmCompras.txtRuc.Text = Trim(Me.txtRuc.Text)
            FrmCompras.TxtProveedor.Text = Trim(Me.TxtRazonSocial.Text)
            FrmCompras.txtDireccion.Text = Trim(Me.TxtDireccion1.Text)
            
            
        End If
        
        
        Unload Me
  

End Sub



Private Function get_cliente_uso(ByVal in_dni As String) As Boolean
    strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & in_dni & "'"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 1 Then
       get_cliente_uso = True
    Else
       get_cliente_uso = False
    End If
    
End Function



Private Sub put_estudiante(ByVal in_dni As String)
Dim in_recuperacion As String
Dim in_mama As String
Dim in_papa As String

If Me.chk_mama.Value = 1 Then
   in_mama = "si"
Else
   in_mama = "no"
End If
If Me.chk_papa.Value = 1 Then
   in_papa = "si"
Else
   in_papa = "no"
End If


strCadena = "SELECT * FROM persona_estudiante WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount < 1 Then
   If Me.chkrecuperacion_si.Value = 1 Then
      in_recuperacion = "si"
   Else
      in_recuperacion = "no"
   End If
    
        
    strCadena = "call p_insert_estudiante('" & in_dni & "','" & Me.DtcNivel.BoundText & "','" & Me.DtcGrado.BoundText & "','" & Trim(Me.txtieprocedencia.Text) & "','" & Me.txtpromovido.Text & "','" & Me.txttercio.Text & "','" & Me.txthabilidad.Text & "','" & in_recuperacion & "','" & Me.DtcTiponacimiento.BoundText & "','" & Trim(Me.txtenfermedades.Text) & "','" & Trim(Me.txtvacunas.Text) & "','" & Trim(Me.txtalergias.Text) & "','" & Me.Dtcseguro.BoundText & "','" & in_papa & "','" & in_mama & "','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
Else
     If Me.chkrecuperacion_si.Value = 1 Then
      in_recuperacion = "si"
   Else
      in_recuperacion = "no"
   End If
    strCadena = "call p_update_estudiante('" & in_dni & "','" & Me.DtcNivel.BoundText & "','" & Me.DtcGrado.BoundText & "','" & Trim(Me.txtieprocedencia.Text) & "','" & Me.txtpromovido.Text & "','" & Me.txttercio.Text & "','" & Me.txthabilidad.Text & "','" & in_recuperacion & "','" & Me.DtcTiponacimiento.BoundText & "','" & Trim(Me.txtenfermedades.Text) & "','" & Trim(Me.txtvacunas.Text) & "','" & Trim(Me.txtalergias.Text) & "','" & Me.Dtcseguro.BoundText & "','" & in_papa & "','" & in_mama & "','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If


If Me.chkseguro.Value = 1 Then
   If Me.chk_estadoseguro.Value = 1 Then
      in_estado_seguro = "si"
   Else
      in_estado_seguro = "no"
   End If
    strCadena = "CALL p_put_seguro_persona('" & in_dni & "','" & Me.Dtcseguro.BoundText & "','" & Trim(Me.txtPoliza.Text) & "','" & Format(Me.dtpExpedicion.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpExpiracion.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & in_estado_seguro & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If
End Sub


Private Sub load_estudiante(ByVal in_dni As String, ByVal in_matricula As String)
strCadena = "SELECT * FROM persona_estudiante WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Me.DtcNivel.BoundText = rstK("id_nivel")
   Me.txtieprocedencia.Text = rstK("procedencia")
   Me.DtcPeriodo.BoundText = rstK("id_periodo")
   Me.txtpromovido.Text = rstK("promovido")
   Me.txttercio.Text = rstK("tercio_estudiantil")
   Me.txthabilidad.Text = rstK("habilidad")
   Me.txtalergias.Text = rstK("alergias")
   Me.txtenfermedades.Text = rstK("enfermedades")
   Me.txtvacunas.Text = rstK("vacunas")
   'Me.DtcGrado.BoundText = rstK("id_grado")
  
   Me.DtcTiponacimiento.BoundText = rstK("id_tipo_nacimiento")
   If rstK("requiere_recuperacion") = "si" Then
      Me.chkrecuperacion_si.Value = 1
   Else
      Me.chkrecuperacion_no.Value = 1
   End If
   If rstK("id_seguro") <> "00000" Then
      Me.chkseguro.Value = 1
      Me.Dtcseguro.BoundText = rstK("id_seguro")
      Call load_seguro(in_dni, Me.Dtcseguro.BoundText)
    Else
      Me.chkseguro.Value = 0
   End If
    
End If



If Val(in_matricula) > 0 Then
    strCadena = "SELECT * FROM college_matricula WHERE dni='" & Trim(Me.txtRuc.Text) & "' and  id_matricula='" & Val(in_matricula) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
       Me.DtcPeriodo.BoundText = rstK("id_periodo")
       Me.DtcNivel.BoundText = rstK("id_nivel")
       Call get_grado(Me.DtcNivel.BoundText)
       Me.DtcGrado.BoundText = rstK("id_grado")
       Me.DtcMatricula.BoundText = rstK("codigo_matricula")
       Me.DtcPension.BoundText = rstK("codigo_pension")
       
       
        Me.DtcEstado.BoundText = rstK("id_estado")
       If rstK("beca_completa") = "si" Then
          Me.chk_beca.Value = 1
       End If
       
       If rstK("media_beca") = "si" Then
           Me.chk_media_beca.Value = 1
       End If
        
    End If
End If


End Sub
Private Sub load_seguro(ByVal in_dni As String, ByVal in_seguro As String)
strCadena = "SELECT * FROM persona_seguro where dni='" & in_dni & "' and id_detalle='" & in_seguro & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   Me.txtPoliza.Text = rstL("numero")
   Me.dtpExpedicion.Value = Format(rstL("expedicion"), "dd-mm-YYYY")
   Me.DtpExpiracion.Value = Format(rstL("expiracion"), "dd-mm-YYYY")
   Me.frmseguro.Visible = True
End If
End Sub
Private Sub quitar_almacen(ByVal dni As String)
strCadena = "DELETE FROM almacen WHERE id_responsable='" & dni & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
End Sub
Private Sub persona_almacen(ByVal dni As String)
Dim strAlm As String
        strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "' AND id_responsable='" & Trim(dni) & "' ORDER BY id_alm DESC"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount < 1 Then
        If rstT.RecordCount > 0 Then
            rstT.MoveFirst
            strAlm = formato_item(Val(rstT("id_alm")) + 1, 5)
        Else
            
    strAlm = formato_item(Val(LastRegistro("almacen", "id_alm")) + 1, 5)
        End If
        strCadena = "INSERT INTO almacen (id_alm,descripcion,direccion,id_responsable,ruc)VALUES ('" & strAlm & "','" & Me.TxtRazonSocial.Text & "','" & Me.TxtDireccion1.Text & "','" & Trim(dni) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        
        strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,stock,ruc) VALUES ('" & strAlm & "','" & rst("id_producto") & "','0','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
             
            rst.MoveNext
        Next i
       Set rst = Nothing
    End If
    End If
End Sub












Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_pago_mensual_Click()

If Me.chk_pago_mensual.Value = 1 Then
   
   Me.lblprorroga.Visible = True
   Me.txtProrroga.Visible = True
   Me.txtUltimodia_pago.Visible = True
   Me.Label8(1).Visible = True
   Me.chk_ultimo_dia_mes.Visible = True
Else
  
  Me.lblprorroga.Visible = False
  Me.txtProrroga.Visible = False
  Me.txtUltimodia_pago.Visible = False
  Me.Label8(1).Visible = False
  Me.chk_ultimo_dia_mes.Visible = False
End If

End Sub

Private Sub chk_seguro_transporte_Click()
If Me.chk_seguro_transporte.Value = 1 Then
   Call load_seguro_lista
End If

End Sub



Private Sub chkDescuento_Click()
If (Me.chkDescuento.Value = 1) Then
    Me.txtDescuento.Visible = True
Else
    Me.txtDescuento.Visible = False
End If
End Sub

Private Sub chkEmpresa_Click()
If Me.chkEmpresa.Value = True Then
    Me.txtRucEmpresa.Visible = True
    
    Me.txtMaximoCredito.Visible = False
    
Else
  
   Me.txtRucEmpresa.Visible = False
   
   
End If
End Sub

Private Sub chkmatricula_Click()

End Sub

Private Sub ChkMaximoCredito_Click()
If Me.ChkMaximoCredito.Value = True Then
    Me.txtMaximoCredito.Visible = True
    Me.txtRucEmpresa.Visible = False
    Me.LblEmpresa.Visible = False
Else
    Me.txtMaximoCredito.Visible = False
End If
End Sub





Private Sub chkseguro_Click()
If Me.chkseguro.Value = 1 Then
   Me.frmseguro.Visible = True
Else
   Me.frmseguro.Visible = False
End If
End Sub



Private Sub cmdActualizar_password_Click()

Me.frmpassword.Visible = True



End Sub

Private Sub cmdagregar_Click()
Dim razon As String
If Me.txtDni.Text <> "" Then
    strCadena = "SELECT * FROM persona_accidentes WHERE dni_familia='" & Trim(Me.txtDni.Text) & "' and id_parentesco='" & Me.dtcparentesco.BoundText & "' and dni='" & Trim(Me.txtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "UPDATE persona_accidentes SET telefono='" & Me.txtTelefono.Text & "',id_parentesco='" & Me.dtcparentesco.BoundText & "',id_ocupacion='" & Me.DtcOcupacion.BoundText & "',id_grado='" & Me.DtcGradoinstruccion.BoundText & "',direccion='" & Trim(Me.txtdireccion_familia.Text) & "' WHERE dni_familia='" & Me.txtDni.Text & "' AND dni='" & Me.txtRuc.Text & "' and id_parentesco='" & Me.dtcparentesco.BoundText & "'"
        CnBd.Execute (strCadena)
        
    Else
        
        strCadena = "INSERT INTO persona_accidentes(dni,dni_familia,id_parentesco,telefono,direccion,id_ocupacion,id_grado)VALUES('" & Me.txtRuc.Text & "','" & Me.txtDni.Text & "','" & Me.dtcparentesco.BoundText & "','" & Me.txtTelefono.Text & "','" & Trim(Me.txtdireccion_familia.Text) & "','" & Me.DtcOcupacion.BoundText & "','" & Me.DtcGradoinstruccion.BoundText & "')"
        CnBd.Execute (strCadena)
        
         
        strCadena = "select * from persona where dni='" & Trim(Me.txtDni.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
        razon = Me.TxtFnombers.Text + Space(1) + Me.TxtFpaterno.Text + Space(1) + Me.TxtFmaterno.Text
        strCadena = "P_insert_persona('" & Trim(Me.txtDni.Text) & "','" & Me.TxtFpaterno.Text & "','" & Me.TxtFmaterno.Text & "','" & Me.TxtFnombers.Text & "','" & Trim(razon) & "','" & Trim(Me.txtdireccion_familia.Text) & "','" & Me.txtTelefono.Text & "','--','no','no','no','no','no','0','')"
        CnBd.Execute (strCadena)
        
         
        End If
    End If
    
    Call llenarFamiliares(Me.HfgFamiliares)
    
End If
End Sub

Private Sub cmdagregar_direccion_Click()


If Trim(Me.lblCodigoUbigeoSunat2.Caption) <> "" Then
    strCadena = "call ADM_direcccion_persona('" & Me.txt_id_direccion.Text & "','" & Replace(Me.txtnuevadireccion.Text, "'", "") & "','" & Trim(Me.txtRuc.Text) & "','" & Me.lblCodigoUbigeoSunat2.Caption & "')"
    CnBd.Execute (strCadena)
Else
        MsgBox "Ingrese un Ubigeo Correcto.", vbInformation
        Exit Sub
End If

    

Call Me.llenar_direccion(Me.hfdireccion, Trim(Me.txtRuc.Text))
Me.frmdireccion.Visible = False
End Sub

Private Sub cmdagregarcuenta_Click()
If Me.txtnumerocuenta.Text <> "" And Trim(Me.DtcBanco.BoundText) <> "" And Trim(Me.DtcMoneda.BoundText) Then
    strCadena = "INSERT INTO persona_cuentabancaria(dni,id_banco,id_moneda,cuenta)VALUES('" & Trim(Me.txtRuc.Text) & "','" & Me.DtcBanco.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Trim(Me.txtnumerocuenta.Text) & "')"
    CnBd.Execute (strCadena)
     
     
    
    Call llenarCuentas(Me.HfCuentas, Trim(Me.txtRuc.Text))
End If
End Sub

Private Sub cmdCerrar_chofer_Click()
Me.frmconductores.Visible = False
End Sub

Private Sub cmdcerrardetalle_Click()
frmplandetalle.Visible = False
End Sub

Private Sub cmdcerrardireccion_Click()
Me.frmdireccion.Visible = False
End Sub

Private Sub cmdCerrarMarca_Click()
Me.frmmarca.Visible = False
End Sub

Private Sub cmdcerrarseguro_Click()
Me.frmseguro.Visible = False
End Sub

Private Sub cmdcertificado_Click()

End Sub

Private Sub cmdconfirmar_pass_Click()
If Trim(Me.TxtPassword.Text) = Trim(Me.txtpassword_confirmar.Text) And Trim(Me.TxtPassword.Text) <> "" Then
   strCadena = "UPDATE entidad_empresa SET password='" & Trim(Me.TxtPassword.Text) & "',passwordaccesso='" & Trim(Me.TxtPassword.Text) & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' and id_empresa='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Me.frmpassword.Visible = False
Else
    MsgBox "Password Ingresados NO COINCIDEN", vbInformation
End If
End Sub

Private Sub cmdConSUNAT_Click()
 Call precionar
End Sub
Public Sub precionar()
  
  If Len(Trim(Me.txtRuc.Text)) = 8 Then
        strCadena = "SELECT dni FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Call LLENA_NC(Trim(Me.txtRuc.Text))
        Else
             If get_dni_reniec_iii(Trim(Me.txtRuc.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                Call LLENA_NC(Trim(Me.txtRuc.Text))
            End If
        End If
        
        
  End If
    
End Sub


Private Sub cmdeditarTelefono_Click()
Me.lblidtelefono.Caption = Val(Me.HfTelefonos.TextMatrix(Me.HfTelefonos.Row, 0))
Me.TxtFono.Text = Trim(Me.HfTelefonos.TextMatrix(Me.HfTelefonos.Row, 2))
End Sub

Private Sub cmdEliminaMarca_Click()
If Val(Me.HfMarcas.TextMatrix(Me.HfMarcas.Row, 0)) > 0 Then
   strCadena = "DELETE FROM persona_transporte where id='" & Val(Me.HfMarcas.TextMatrix(Me.HfMarcas.Row, 0)) & "'"
   CnBd.Execute (strCadena)
   Call Me.llenar_marcas(Me.HfMarcas, Trim(Me.txtRuc.Text))

End If
End Sub

Private Sub cmdeliminar_direccion_Click()
If Val(Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0)) > 0 Then
    If MsgBox("Esta seguro de Eliminar esta direccion", vbYesNo, KEY_EMPRESA) = vbYes Then
   strCadena = "p_delete_direccion('" & Val(Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0)) & "')"
   CnBd.Execute (strCadena)
   Call Me.llenar_direccion(Me.hfdireccion, Trim(Me.txtRuc.Text))
End If
End If
End Sub

Private Sub cmdeliminar_item_Click()
Call delete_parentesco(Val(Me.HfgFamiliares.TextMatrix(Me.HfgFamiliares.Row, 0)))

Call llenarFamiliares(Me.HfgFamiliares)

End Sub

Private Sub delete_parentesco(ByVal in_id As String)

strCadena = "DELETE FROM persona_accidentes WHERE id='" & Val(Me.HfgFamiliares.TextMatrix(Me.HfgFamiliares.Row, 0)) & "'"
CnBd.Execute (strCadena)

End Sub

Private Sub cmdEliminarchofer_Click()
If Val(Me.HfChofer.TextMatrix(Me.HfChofer.Row, 0)) > 0 Then
   strCadena = "DELETE FROM persona_chofer where id_persona='" & Trim(Me.txtRuc.Text) & "' and dni='" & Trim(Me.HfChofer.TextMatrix(Me.HfChofer.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Call Me.llenar_chofer(Me.HfChofer, Trim(Me.txtRuc.Text))
End If
End Sub

Private Sub cmdeliminarcuenta_Click()
If Val(Me.HfCuentas.TextMatrix(Me.HfCuentas.Row, 1)) > 0 And Trim(Me.txtRuc.Text) <> "" Then
    strCadena = "DELETE FROM persona_cuentabancaria WHERE cuenta='" & Me.HfCuentas.TextMatrix(Me.HfCuentas.Row, 1) & "' AND dni='" & Trim(Me.txtRuc.Text) & "'"
    CnBd.Execute (strCadena)
     
    Call llenarCuentas(Me.HfCuentas, Trim(Me.txtRuc.Text))
End If
End Sub

Private Sub cmdGrabarFoto_Click()

End Sub

Private Sub cmdeliminarplan_Click()
Call disabled_form(FrmPersona)
Call disabled_form(Me)

Procedencia = Eliminar
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdModificaMarca_Click()

End Sub

Private Sub cmdmodificar_direccion_Click()
Me.txt_id_direccion.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0)
strCadena = "SELECT * FROM persona_direccion WHERE id_direccion='" & Val(Me.txt_id_direccion.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtnuevadireccion.Text = rst("direccion")
   Me.lblCodigoUbigeoSunat2.Caption = rst("codigo_ubigeo_sunat")
   'strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito where id_distrito='" & rst("id_distrito") & "'"
   'Call ConfiguraRstT(strCadena)
   'Call LlenaDataComboT(Me.DtcDistrito2)
   'Call buscar_ubigueo
   Call Me.get_ubigeo_sunat2(rst("codigo_ubigeo_sunat"))
End If
Me.frmdireccion.Visible = True

Me.cmdagregar_direccion.Enabled = True
End Sub

Private Sub load_servicio()



strCadena = "SELECT id_plan as Codigo,descripcion as Descripcion FROM plan_servicio WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPlanServicio)

'cargar Medio de Contacto
strCadena = "SELECT id_medio as Codigo,descripcion as Descripcion FROM empresa_medio_contacto"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMedioContacto)

'Cargar Estado de Empresa
strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM empresa_estado"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoEmpresa)


strCadena = "SELECT id_prioridad as Codigo,descripcion as Descripcion FROM proyecto_prioridad"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcprioridad)



strCadena = "SELECT * FROM persona_plan_servicio_ii where dni='" & Trim(Me.txtRuc.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.DtcPlanServicio.BoundText = rst("id_plan")
    Me.txt_precio_plan.Text = get_precio_producto(rst("id_producto"), rst("id_alm"))
    Me.DtcSucursal.BoundText = rst("id_alm")
    Me.txt_precio_plan.Text = rst("monto")
    Me.DtpFechaSuscripcion.Value = rst("fecha_suscripcion")
    If rst("pago_anual") = "si" Then
       Me.chk_anual.Value = 1
    End If
    
    If rst("pago_6meses") = "si" Then
        Me.chk_6meses.Value = 1
    End If
    
    If rst("pago_3meses") = "si" Then
       Me.chk_3meses.Value = 1
    End If
    
    
    Me.DtcEstadoEmpresa.BoundText = rst("id_estado")
    Me.DtcMedioContacto.BoundText = rst("id_medio_contacto")
    Me.dtcprioridad.BoundText = rst("id_prioridad")
    
    
    
    If rst("pago_mensual") = "si" Then
       Me.chk_pago_mensual.Value = 1
       Me.lblprorroga.Visible = True
       Me.txtProrroga.Visible = True
       Me.txtProrroga.Text = rst("mora_dias")
       If rst("ultimo_dia_mes") = "si" Then
          Me.txtUltimodia_pago.Text = ""
          Me.chk_ultimo_dia_mes.Value = 1
       Else
          Me.txtUltimodia_pago.Text = rst("dia_pago")
          Me.chk_ultimo_dia_mes.Value = 0
       End If
    Else
       Me.lblprorroga.Visible = False
       Me.txtProrroga.Visible = False
       Me.txtProrroga.Text = 0
    End If
   
End If

strCadena = "SELECT activacion_corte FROM entidad_empresa WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If rst("activacion_corte") = "si" Then
        Me.chk_activacion_corte.Value = 1
    Else
        Me.chk_activacion_corte.Value = 0
    End If
End If

 Me.frmplandetalle.Visible = True
End Sub

Private Sub cmdmodificarplan_Click()
Call load_servicio
End Sub

Private Sub cmdnuevo_direccion_Click()
Me.txt_id_direccion.Text = ""
Me.frmdireccion.Visible = True
Call Resalta(Me.txtnuevadireccion)
Me.cmdagregar_direccion.Enabled = True
End Sub

Private Sub cmdNuevoChofer_Click()
Me.frmconductores.Visible = True
Me.txtchofer.Text = ""

Me.txtlicenciaTransporte.Text = ""
End Sub

Private Sub cmdNuevoMarca_Click()

   Me.TxtIdMarca.Text = ""
   Me.TxtMarca.Text = ""
   Me.TxtPlaca.Text = ""
   Me.txtcertificado.Text = ""
   Me.frmmarca.Visible = True

End Sub

Private Sub cmdnuevoplan_Click()

 strCadena = "SELECT id_plan as Codigo,CONCAT(descripcion,'- [',round(precio_venta,2),' ]') as Descripcion FROM view_plan WHERE id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 Call LlenaDataCombo(Me.DtcPlanServicio)
 
 


'cargar Medio de Contacto
strCadena = "SELECT id_medio as Codigo,descripcion as Descripcion FROM empresa_medio_contacto"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMedioContacto)

'Cargar Estado de Empresa
strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM empresa_estado"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoEmpresa)


strCadena = "SELECT id_prioridad as Codigo,descripcion as Descripcion FROM proyecto_prioridad"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcprioridad)



 
Me.txt_precio_plan.Text = ""
Me.txtUltimodia_pago.Text = 1

Me.chk_pago_mensual.Value = 0
Me.chk_6meses.Value = 0
Me.chk_3meses.Value = 0
Me.frmplandetalle.Visible = True

End Sub

Private Sub cmdprocesar_chofer_Click()
If Trim(Me.txtchofer.Text) <> "" And Trim(Me.txtlicenciaTransporte.Text) <> "" Then
    
    If get_persona_existe(Trim(Me.txtchofer.Text)) = False Then
       strCadena = "INSERT INTO persona (dni,nombre_completo)VALUES('" & Trim(Me.txtchofer.Text) & "','" & Trim(Me.lblchofer.Text) & "')"
       CnBd.Execute (strCadena)
    End If
    
    strCadena = "sp_chofer('" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtchofer.Text) & "','" & Trim(Me.txtlicenciaTransporte.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call Me.llenar_chofer(Me.HfChofer, Trim(Me.txtRuc.Text))
    Me.frmconductores.Visible = False
End If
End Sub

Private Sub cmdProcesar_Click()
On Error GoTo error
  
      Call Save
       Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub cmdprocesarMarca_Click()

If Trim(Me.TxtMarca.Text) <> "" And Trim(Me.TxtPlaca.Text) <> "" Then
    strCadena = "call sp_marca_placa_v2('" & Val(Me.TxtIdMarca.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtMarca.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.txtcertificado.Text) & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtMotor.Text) & "','" & Trim(Me.txtcolor.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call Me.llenar_marcas(Me.HfMarcas, Trim(Me.txtRuc.Text))
    Me.frmmarca.Visible = False
End If


End Sub

Private Sub cmdprocesarplan_Click()

If Me.chk_anual.Value = 1 Then
   in_pago_anual = "si"
Else
   in_pago_anual = "no"
End If

If Me.chk_6meses.Value = 1 Then
    in_pago_6meses = "si"
Else
    in_pago_6meses = "no"
End If

If Me.chk_3meses.Value = 1 Then
   in_pago_3meses = "si"
Else
   in_pago_3meses = "no"
End If

If Me.chk_pago_mensual.Value = 1 Then
    in_pago_mensual = "si"
Else
    in_pago_mensual = "no"
    Me.txtProrroga.Text = 0
End If

If Me.chk_ultimo_dia_mes.Value = 1 Then
    in_ultimo_dia = "si"
Else
    in_ultimo_dia = "no"
End If

If Me.chk_activacion_corte.Value = 1 Then
    in_activacion_corte = "si"
Else
    in_activacion_corte = "no"
End If

If Trim(Me.DtcPlanServicio.BoundText) > 0 Then
    strCadena = "call put_plan_servicio_14('" & Me.DtcPlanServicio.BoundText & "','" & Trim(Me.txtRuc.Text) & "','" & in_pago_anual & "','" & in_pago_6meses & "','" & in_pago_3meses & "','" & in_pago_mensual & "','" & KEY_USUARIO & "','" & Me.DtcSucursal.BoundText & "','" & Val(Me.txtUltimodia_pago.Text) & "','" & Format(Me.DtpFechaSuscripcion.Value, "YYYY-mm-dd") & "','" & Val(Me.txtProrroga.Text) & "','" & KEY_MORA_MONTO & "','" & Val(Me.txt_precio_plan.Text) & "','" & in_ultimo_dia & "','" & Me.DtcEstadoEmpresa.BoundText & "','" & Me.DtcMedioContacto.BoundText & "','" & Me.dtcprioridad.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If


strCadena = "UPDATE entidad_empresa SET activacion_corte='" & in_activacion_corte & "',fecha_corte='" & Format(Me.DtpFechaCorte.Value, "YYYY-mm-dd") & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' and id_empresa='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "UPDATE entidad_parametros SET caducidad='" & Format(Me.DtpFechaCorte.Value, "YYYY-mm-dd") & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
CnBd.Execute (strCadena)

Call llenar_plan_servicio(Hfplanservicio, Trim(Me.txtRuc.Text))
frmplandetalle.Visible = False

If KEY_RUBRO = "00026" Then
   ' Call put_insertar_plantilla
End If




End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTransporte_Click()
If Me.frmtransporte.Visible = True Then
    Me.frmtransporte.Visible = False
Else
    Call Me.llenar_marcas(Me.HfMarcas, Trim(Me.txtRuc.Text))
    Call Me.llenar_chofer(Me.HfChofer, Trim(Me.txtRuc.Text))
    Me.frmtransporte.Visible = True
End If
End Sub

Private Sub cmdVisualizar_Click()
If Me.txtRuc.Text <> "" Then
    Call LLENA_NC(Trim(Me.txtRuc.Text))
End If
End Sub

Private Sub Command1_Click()
Dim cPersona As Double
    If (Me.txtRuc.Text) <> "" Then
        If Trim(Me.TxtFono.Text) = "" Then
           MsgBox "Ingrese un TELEFONO Valido", vbInformation, "Mensaje para el Usuario"
           Call Resalta(Me.TxtFono)
           Exit Sub
        End If
        
        If Val(Me.lblidtelefono.Caption) > 0 Then
            
            strCadena = "UPDATE persona_telefono SET telefono='" & Trim(Me.TxtFono.Text) & "' WHERE id_telefono='" & Val(Me.lblidtelefono.Caption) & "'"
        Else
            strCadena = "INSERT INTO persona_telefono (dni,telefono,id_cargo)VALUES ('" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtFono.Text) & "','" & Me.DtcArea.BoundText & "')"
        End If
        CnBd.Execute (strCadena)
        Me.lblidtelefono.Caption = 0
        
        strCadena = "UPDATE persona SET celular='" & Trim(Me.TxtFono.Text) & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
        CnBd.Execute (strCadena)
        If Trim(Me.txtRuc.Text) = KEY_RUC Then
            Call get_telefonos(Trim(Me.txtRuc.Text))
        End If
         
    Else
        MsgBox "Ingrese un Ruc/DNI", vbInformation, "Mensaje para el Usuario"
        Call Resalta(Me.txtRuc)
        Exit Sub
    End If
    
    Me.TxtFono.Text = ""
    Call Resalta(Me.TxtFono)
    Call LlenarTelefonos(Me.HfTelefonos, Trim(Me.txtRuc.Text))

End Sub
Public Sub LlenarTelefonos(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM view_telefono WHERE dni='" & cPersona & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 2000
            Grilla.ColWidth(2) = 1500
        Next
        cabecera = "codigo" & vbTab & "AREA" & vbTab & "TELEFONO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_telefono") & vbTab & rst("descripcion") & vbTab & rst("telefono")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
     
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub llenar_direccion(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM persona_direccion WHERE dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdmodificar_direccion.Enabled = False
    Me.cmdeliminar_direccion.Enabled = False
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 500
            Grilla.ColWidth(2) = 4500
            
        Next
        cabecera = "COD" & vbTab & "COD" & vbTab & "DIRECCION"
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_direccion") & vbTab & Format(i + 1, "0000") & vbTab & rst("direccion")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
     Me.cmdmodificar_direccion.Enabled = False
    Me.cmdeliminar_direccion.Enabled = False
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenar_plan_servicio(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM view_plan_servicio_persona WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdmodificarplan.Enabled = False
    Me.cmdeliminarplan.Enabled = False
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 2500
            Grilla.ColWidth(2) = 700
            Grilla.ColWidth(3) = 700
            Grilla.ColWidth(4) = 700
            Grilla.ColWidth(5) = 700
            Grilla.ColWidth(6) = 700
            Grilla.ColWidth(7) = 1000
            Grilla.ColWidth(8) = 1200
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "ANUAL" & vbTab & "6 MESES" & vbTab & "3 MESES" & vbTab & "MENSUAL" & vbTab & "MONTO" & vbTab & "F.INICIO" & vbTab & "F.PAGO"
        Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If rst("pago_anual") = "si" Then
               in_anual = "  X  "
            Else
               in_anual = "  "
            End If
            
            If rst("pago_6meses") = "si" Then
               in_6meses = "  X  "
            Else
               in_6meses = "  "
            End If
            
            If rst("pago_3meses") = "si" Then
               in_3meses = "  X  "
            Else
               in_3meses = "  "
            End If
            If rst("pago_mensual") = "si" Then
               in_mensual = "  X  "
            Else
               in_mensual = "  "
            End If
            
            If rst("ultimo_dia_mes") = "si" Then
               
               fecha_pago = Format(DateSerial(Year(KEY_FECHA), Month(KEY_FECHA) + 1, 0), "dd-mm-YYYY")
            Else
               fecha_pago = "PAGA LOS: " & Format(rst("dia_pago"), "00")
            End If
            
            
            
            
            Fila = rst("id") & vbTab & rst("descripcion") & vbTab & in_anual & vbTab & in_6meses & vbTab & in_3meses & vbTab & in_mensual & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(rst("fecha_suscripcion"), "dd-mm-YYYY") & vbTab & fecha_pago
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
    
        
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenarCuentas(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
strCadena = "SELECT abreviatura,M.descripcion,cuenta FROM banco B,moneda M,persona_cuentabancaria C WHERE B.id_banco=C.id_banco AND C.id_moneda=M.id_moneda AND C.dni='" & cPersona & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 1000
            Grilla.ColWidth(1) = 1500
            Grilla.ColWidth(2) = 1500
        Next
        cabecera = "BANCO" & vbTab & "CUENTA" & vbTab & "MONEDA"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = Fila & rstT("abreviatura") & vbTab & rstT("cuenta") & vbTab & UCase(rstT("descripcion"))
            Grilla.AddItem Fila
            Fila = ""
            rstT.MoveNext
        Next i
 Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub

Public Sub llenar_marcas(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
strCadena = "SELECT * FROM persona_transporte WHERE id_persona='" & cPersona & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 500
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 1200
            Grilla.ColWidth(3) = 500
            Grilla.ColWidth(4) = 1200
            Grilla.ColWidth(5) = 1200
        Next
        cabecera = "CODIGO" & vbTab & "MARCA" & vbTab & "PLACA" & vbTab & "CERTIFICADO" & vbTab & "SERIE" & vbTab & "MOTOR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = Format(rstT("id"), "0000") & vbTab & rstT("marca") & vbTab & rstT("placa") & vbTab & rstT("certificado") & vbTab & rstT("serie") & vbTab & rstT("motor")
            Grilla.AddItem Fila
            rstT.MoveNext
        Next i
 Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub

Public Sub llenar_ubigeo(ByVal Grilla As MSHFlexGrid, ByVal in_busqueda As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_ubigeo_sunat WHERE ubigeo LIKE '%" & Replace(in_busqueda, "'", "") & "%'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 800
            Grilla.ColWidth(1) = 6700
            
        Next
        cabecera = "CODIGO" & vbTab & "UBIGEO"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("cod_ubigeo_sunat") & vbTab & rstT("ubigeo")
            Grilla.AddItem Fila
            rstT.MoveNext
        Next i
 Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub

Public Sub llenar_chofer(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_chofer_empresa WHERE id_persona='" & cPersona & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 1000
            Grilla.ColWidth(1) = 3500
            Grilla.ColWidth(2) = 1500
        Next
        cabecera = "DNI" & vbTab & "NOMBRE COMPLETO" & vbTab & "LICENCIA"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("dni") & vbTab & rstT("nombre_completo") & vbTab & rstT("licencia")
            Grilla.AddItem Fila
            rstT.MoveNext
        Next i
 Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub

Private Sub Command3_Click()
If Val(Me.HfTelefonos.TextMatrix(Me.HfTelefonos.Row, 0)) > 0 Then
    If Trim(Me.txtRuc.Text) <> "" Then
        strCadena = "DELETE FROM persona_telefono WHERE id_telefono='" & Val(Me.HfTelefonos.TextMatrix(Me.HfTelefonos.Row, 0)) & "'"
        CnBd.Execute (strCadena)
        Call LlenarTelefonos(Me.HfTelefonos, Me.txtRuc.Text)
        Exit Sub
 End If
End If
    
        
End Sub





Private Sub cSunat_DatosObtenidos()
    Dim Confirm As Integer
    Dim in_razon_social As String
    Dim nombres() As String
    Dim telefonos() As String
    If cSunat.EmRazSocial <> vbNullString Then
        If TxtRazonSocial.Text = vbNullString Then
            TxtRazonSocial.Text = cSunat.EmRazSocial
            in_razon_social = Trim(Me.TxtRazonSocial.Text)
            
            nombres = Split(Trim(TxtRazonSocial.Text), " ")
            Me.txtPaterno.Text = nombres(0)
            'Me.txtMaterno.Text = nombres(1)
            
            If UBound(nombres()) > 3 Then
                Me.txtNombre.Text = nombres(2) & Space(1) & nombres(3)
            Else
                If UBound(nombres()) > 1 Then
                    Me.txtNombre.Text = nombres(2)
                End If
                If UBound(nombres()) >= 3 Then
                    Me.txtNombre.Text = nombres(2) & Space(1) & nombres(3)
                End If
                
            End If
            Me.TxtRazonSocial.Text = in_razon_social
            Me.TxtDireccion1.Text = cSunat.EmDireccion
            
            If Me.TxtDireccion1.Text = "-" Then
                Me.TxtDireccion1.Text = KEY_DIR_PUBLIC
            End If
            
            
            
            If Len(Trim(Me.txtRuc.Text)) = 11 Then
                If MsgBox("Desea registrar con RUC  ", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
                    Me.txtRuc.Text = Trim(Me.txtRuc.Text)
                Else
                    Me.txtRuc.Text = Mid(Trim(Me.txtRuc.Text), 3, 8)
                End If
                    Me.LblCodPersona.Caption = Trim(Me.txtRuc.Text)
            End If
            
            
        Else
            If TxtRazonSocial.Text <> cSunat.EmRazSocial Then
                Confirm = MsgBox("La raz�n social que tiene almacenado el sistema no coincide con la informacion de SUNAT, �desea actualizar?", vbYesNo, "Confirmar actualizaci�n")
                If Confirm = vbYes Then
                    TxtRazonSocial.Text = cSunat.EmRazSocial
                     Me.TxtDireccion1.Text = cSunat.EmDireccion
                     
                End If
            End If
            
          
           ' If txtNomComercial.Text = vbNullString Then
           ' txtNomComercial.Text = cSunat.EmNomComercial
           ' End If
            
        End If
        
        'cmdGuardar.SetFocus
         
        Else
         If Len(Trim(Me.txtRuc.Text)) > 8 And Trim(Me.TxtRazonSocial.Text) = "" Then
            'Me.txtRUC.Text = Mid(Trim(Me.txtRUC.Text), 3, 8)
        
        End If
    End If
    
End Sub

Private Sub cSunat_ErrorEnObtencion()
    MsgBox cSunat.ErrConSunat
End Sub

'============================================================


Private Sub CmdFoto_Click()
FrmCapturarImagen.Show
'On Error GoTo finish
'Me.CommonDialog1.Filter = "*.Jpg"
'Me.CommonDialog1.ShowOpen
'Me.Image1.Picture = LoadPicture(Me.CommonDialog1.FileName)
'img = Me.CommonDialog1.FileName
'Exit Sub
'finish: MsgBox "La Imagen que Intenta Subir tiene que ser .JPG", vbInformation, "Imagen no Compatible"

End Sub

Private Sub DtcDistrito_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Me.DtcDistrito.BoundText <> "" Then
'        strCadena = "SELECT id_provincia FROM distrito WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
'        Call ConfiguraTemporal(strCadena)
'        If rstTemporal.RecordCount > 0 Then
'            Me.lblprovincia(0).Visible = True
'            Me.DtcProvincia.Visible = True
'            strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_provincia='" & rstTemporal("id_provincia") & "'"
'            Call ConfiguraRst(strCadena)
'            Call LlenaDataCombo(Me.DtcProvincia)
'            Me.DtcProvincia.Enabled = False
'        End If
'    End If
'End If
End Sub

Private Sub DtcDistrito2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_ubigueo
End If
End Sub
Private Sub buscar_ubigueo()
'If Me.DtcDistrito2.BoundText <> "" Then
'        strCadena = "SELECT id_provincia FROM distrito WHERE id_distrito='" & Me.DtcDistrito2.BoundText & "'"
'        Call ConfiguraTemporal(strCadena)
'        If rstTemporal.RecordCount > 0 Then
'            strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_provincia='" & rstTemporal("id_provincia") & "'"
'            Call ConfiguraRst(strCadena)
'            Call LlenaDataCombo(Me.DtcProvincia2)
'            Me.DtcProvincia2.Enabled = False
'        End If
'    End If
End Sub
Private Sub Dtcmes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtAnio)
End If
End Sub

Private Sub get_grado(ByVal in_nivel As String)
strCadena = "SELECT id_grado as Codigo,descripcion as Descripcion FROM nivel_educativo_grado WHERE id_nivel='" & in_nivel & "' and id_periodo='" & Me.DtcPeriodo.BoundText & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcGrado)
End Sub

Private Sub DtcNivel_Change()
Call get_grado(Me.DtcNivel.BoundText)
End Sub

Private Sub DtcPeriodo_Change()
Call get_grado(Me.DtcNivel.BoundText)
End Sub

Private Sub DtcPlanServicio_Change()
strCadena = "SELECT * FROM plan_servicio WHERE id_plan='" & Me.DtcPlanServicio.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Me.txt_precio_plan.Text = get_precio_producto(rstT("id_producto"), KEY_ALM)
    End If
End Sub

Private Sub DtcProvincia_Change()
'If Me.DtcProvincia.BoundText <> "" Then
'    strCadena = "SELECT * FROM provincia WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' "
'    Call ConfiguraTemporal(strCadena)
'    If rstTemporal.RecordCount > 0 Then
'        Me.lblDepartamento(1).Visible = True
'        Me.DtcDepartamento.Visible = True
'        strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & rstTemporal("id_departamento") & "'"
'        Call ConfiguraRst(strCadena)
'        Call LlenaDataCombo(Me.DtcDepartamento)
'        Set rst = Nothing
'        Me.DtcDepartamento.Enabled = True
'    End If
'    Set rstTemporal = Nothing
'End If
End Sub


Private Sub DtcZona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Save
End If
End Sub

Private Sub DtcProvincia2_Change()
'If Me.DtcProvincia2.BoundText <> "" Then
'    strCadena = "SELECT * FROM provincia WHERE id_provincia='" & Me.DtcProvincia2.BoundText & "' "
'    Call ConfiguraTemporal(strCadena)
'    If rstTemporal.RecordCount > 0 Then
'        Me.DtcDepartamento2.Visible = True
'        strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & rstTemporal("id_departamento") & "'"
'        Call ConfiguraRst(strCadena)
'        Call LlenaDataCombo(Me.DtcDepartamento2)
'        Set rst = Nothing
'        Me.DtcDepartamento2.Enabled = True
'    End If
'    Set rstTemporal = Nothing
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
    Exit Sub
  End If
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Top = 50
    
    Me.DtpFechaSuscripcion.Value = KEY_FECHA
    
    
    Me.SstKardex.TabVisible(3) = False
    



   
  
  If KEY_RUBRO = "00025" Then
     strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion FROM college_periodo ORDER BY id_periodo DESC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcPeriodo)
     strCadena = "SELECT id_nivel as Codigo,descripcion as Descripcion FROM nivel_educativo ORDER BY id_nivel ASC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcNivel)
     
     
      strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM college_estado ORDER BY id_estado DESC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcEstado)
     
     
     strCadena = "SELECT id_producto as Codigo,CONCAT(nombre_prod,'- [',precio_venta,' ]') as Descripcion FROM view_producto WHERE ruc='" & KEY_RUC & "'"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcMatricula)
     
     strCadena = "SELECT id_producto as Codigo,CONCAT(nombre_prod,'- [',precio_venta,' ]') as Descripcion FROM view_producto WHERE ruc='" & KEY_RUC & "'"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcPension)
     
     Me.SstKardex.TabVisible(3) = True
     strCadena = "SELECT id_tipo_nacimiento as Codigo, descripcion as Descripcion FROM tipo_nacimiento"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcTiponacimiento)
   
     strCadena = "SELECT id_detalle as Codigo,descripcion as Descripcion  FROM seguro_medico_detalle where ruc='" & KEY_RUC & "' "
     Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.Dtcseguro)
     strCadena = "SELECT id_grado as Codigo,descripcion as Descripcion FROM grado_instruccioni order by descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcGradoinstruccion)
  
  strCadena = "SELECT id_ocupacion as Codigo,descripcion as Descripcion FROM ocupacion order by id_ocupacion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcOcupacion)
  End If
  
  
  
  
  
  If KEY_RUBRO = "00027" Then
      Me.frmBiblio.Visible = True
     strCadena = "SELECT id_facultad as Codigo,descripcion as Descripcion FROM biblio_facultad ORDER BY id_facultad ASC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcFacultad)
     
      strCadena = "SELECT id_ciclo as Codigo,descripcion as Descripcion FROM biblio_ciclo ORDER BY id_ciclo DESC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcCiclo)
     
     
     strCadena = "SELECT id_tipo_acceso as Codigo,descripcion as Descripcion FROM biblio_tipo_acceso ORDER BY id_tipo_acceso DESC"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcTipoAcceso)
     
  End If
  
  
strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM mes ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)
  
strCadena = "SELECT id_tipo_cliente as Codigo,descripcion as Descripcion FROM tipo_cliente ORDER BY id_tipo_cliente ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(DtcCobertura)
     
strCadena = "SELECT id_parentesco as Codigo,descripcion as Descripcion FROM parentesco ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcparentesco)

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_tipoentidad='0' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSucursal)
Me.DtcSucursal.BoundText = KEY_ALM
 
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE habilitado='si' and  ruc='" & KEY_RUC & "' AND id_personal='si' ORDER BY nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)
  
strCadena = "SELECT id_pais as Codigo,descripcion as Descripcion FROM pais WHERE habilitado='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPais)
  
strCadena = "SELECT id_cargo as Codigo, descripcion as Descripcion FROM persona_cargos WHERE ruc='no' and id_empresa='0' ORDER BY descripcion"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcArea)

strCadena = "SELECT id_frecuencia as Codigo, descripcion as Descripcion FROM frecuencia_visitas"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcFrecuenciaVisita)

strCadena = "SELECT id_tipozona as Codigo, descripcion as Descripcion FROM tipo_zona"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTipoZona)

strCadena = "SELECT id_giro as Codigo, descripcion as Descripcion FROM giro_negocio"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcGiro)


strCadena = "SELECT id_sexo as Codigo, descripcion as Descripcion FROM sexo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcSexo)






If KEY_PAIS = KEY_PERU Then
       Me.DtcPais.BoundText = KEY_PERU
End If


 If KEY_RUBRO <> "00025" Then
 
    'Set objEmpresa = New cbEmpresa
        Set cSunat = New cls_QrySUNAT
        Set cSunat.WebExplorer = wbrInfo
        Set cSunat.WebInet = inetConecta
    
       
        If IdEmpresa = 0 Then
        'ModoInsercion = True
            Me.Caption = "Ficha nueva de empresa"
        'cmdGuardar.Caption = "Guardar"
            If RucExt <> vbNullString Then
                txtRuc.Text = RucExt
                cmdConSUNAT_Click
            Else
                wbrInfo.Navigate "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"
            End If
        Else
       ' ModoInsercion = False
            Me.Caption = "Ficha de empresa N� " & IdEmpresa
        'Call CargarDatosEmpresa(IdEmpresa)
       ' cmdGuardar.Caption = "Actualizar"
        
        txtRuc.Tag = txtRuc.Text
    End If
  End If



Select Case FrmPersona.Procedencia
        Case nuevo
        Me.txtRuc.Enabled = True
        Me.OptSincredito.Value = 1
        Me.ChkCliente.Value = 1
       ' Me.SstKardex.TabVisible(1) = False
        Me.lbldepartamento(1).Visible = False
        Me.DtcDepartamento.Visible = False
        Me.lblprovincia(0).Visible = False
        Me.DtcProvincia.Visible = False
        
        
    Case modificar
      Me.txtRuc.Enabled = False
      
          If LLENA(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0)) = True Then
                Exit Sub
        
      End If
      
      If Len(Me.txtRuc.Text) > 8 Then
      Call precionar
      End If
      
  End Select
  
  
  
  Select Case FrmMatricula.Procedencia
        Case nuevo
        Me.txtRuc.Enabled = True
        Me.OptSincredito.Value = 1
        Me.ChkCliente.Value = 1
       
        Me.lbldepartamento(1).Visible = False
        Me.DtcDepartamento.Visible = False
        Me.lblprovincia(0).Visible = False
        Me.DtcProvincia.Visible = False
        
        
    Case modificar
      Me.txtRuc.Enabled = False
      Me.lbl_id_matricula.Caption = FrmMatricula.HfdPersona.TextMatrix(FrmMatricula.HfdPersona.Row, 0)
          If LLENA(FrmMatricula.HfdPersona.TextMatrix(FrmMatricula.HfdPersona.Row, 1)) = True Then
                Exit Sub
        
      End If
      
      If Len(Me.txtRuc.Text) > 8 Then
      Call precionar
      End If
      
  End Select
  
  
End Sub
Private Sub get_zona(ByVal in_zona As String)
If Val(in_zona) > 0 Then
   strCadena = "SELECT id_zona as Codigo,descripcion_zona as Descripcion  FROM zona WHERE id_zona='" & Val(in_zona) & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcZona)
End If

End Sub
Public Function LLENA(ByVal cPersona As String) As Boolean
'On Error GoTo salir
Dim cDepartamento As String, cProvincia As String, cDistrito As String, cUrbanizacion As Double, cZona As Double
Dim in_seguro As String
Dim in_zona As String
strCadena = "SELECT * FROM view_entidad E WHERE ruc='" & KEY_RUC & "' AND dni = '" & cPersona & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    LLENA = False
    Exit Function
End If
  LLENA = True
  StrCodTabla = rst("dni")
  in_pais = rst("id_pais")
  in_seguro = rst("id_seguro")
  in_zona = rst("id_zona")
  codigo_ubigeo_sunat = rst("codigo_ubigeo_sunat")
  
  Me.DtcSexo.BoundText = rst("sexo")
  
  
  If IsNull(rst("fecha_corte")) = True Then
     Me.DtpFechaCorte.Value = DateAdd("d", 7, KEY_FECHA)
  Else
     Me.DtpFechaCorte.Value = rst("fecha_corte")
  End If
  
  Me.txtObservacion.Text = rst("observacion")
  
  If Len(StrCodTabla) = 8 Then
     Me.SstKardex.TabVisible(5) = True
  Else
   ' Me.SstKardex.TabVisible(5) = False
  End If
  Me.LblCodPersona.Caption = StrCodTabla
  If IsNull(rst("dni")) = True Then
    GoTo sin_doc
  End If
  If rst("habilitado") = "si" Then
     Me.chkhabilitado.Value = 1
  Else
     Me.chkhabilitado.Value = 0
  End If
  
  If rst("afecto_igv") = "si" Then
    Me.chk_igv.Value = 1
  Else
    Me.chk_igv.Value = 0
  End If
  
  If rst("extranjero") = "si" Then
     Me.chk_extranjeria.Value = 1
  End If
  
  
  Me.DtcTipoAcceso.BoundText = rst("id_tipo_acceso")
  Me.DtcFacultad.BoundText = rst("id_facultad")
  Me.DtcCiclo.BoundText = rst("id_ciclo")
  If IsNull(rst("password")) = False Then
    Me.TxtPassword.Text = rst("password")
  End If
  
  Me.txtRuc.Text = rst("dni")
  Me.txtPaterno.Text = UCase(rst("a_paterno"))
  Me.txtMaterno.Text = UCase(rst("a_materno"))
  Me.txtNombre.Text = UCase(rst("nombres"))
  Me.txtpeso.Text = rst("peso")
  Me.txttalla.Text = rst("estatura")
  '  cDepartamento = rst("id_departamento")
  '  cProvincia = rst("id_provincia")
  '  cDistrito = rst("id_distrito")
  
    
  
  Me.TxtLicencia.Text = rst("licencia")
  If Len(rst("dni")) = 8 Then
         Me.LblEntidad.Caption = "Nombre:"
         Me.LblTipoDocumento.Caption = "DNI"
         Call llenarFamiliares(Me.HfgFamiliares)
    Else
        Me.LblEntidad.Caption = "Razon Social:"
        Me.LblTipoDocumento.Caption = "RUC"
    End If
sin_doc:
    Me.TxtRazonSocial.Text = rst("nombre_completo")
        

If (rst("descuento") > 0) Then
    Me.chkDescuento.Value = 1
    Me.txtDescuento.Visible = True
    Me.txtDescuento.Text = rst("descuento")
Else
    Me.chkDescuento.Value = 0
End If


Me.DtcCobertura.BoundText = rst("id_tipo_cliente")



If Val(rst("id_vendedor")) > 0 Then
    
    Me.DtcVendedor.BoundText = rst("id_vendedor")
End If
          strCadena = "SELECT id_banco as Codigo,abreviatura as Descripcion FROM banco ORDER BY abreviatura"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcBanco)
          strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcMoneda)
          Call llenarCuentas(Me.HfCuentas, Trim(Me.LblCodPersona.Caption))
          
  If rst("id_personal") = "si" Then
        Me.ChkPersonal.Value = 1
        StrPersonal = "si"
        Me.SstKardex.TabVisible(1) = True
        Me.DtcVendedor.BoundText = rst("id_vendedor")
  End If
    
    Me.txtdia.Visible = True
    Me.DtcMes.Visible = True
    Me.txtdia.Text = formato_item(rst("id_dia"), 2)
    Me.DtcMes.BoundText = rst("id_mes")
    Me.txtAnio.Text = rst("id_anio")
    'cDepartamento = rst("id_departamento")
    
    Me.TxtDireccion1.Text = rst("direccion")
    
  If Trim(rst("afecto_percepcion")) = "si" Then
    Me.ChkPercepcion.Value = 1
    StrPercepcion = "si"
   Else
    Me.ChkPercepcion.Value = 0
    StrPercepcion = "no"
  End If
  
  If Trim(rst("id_retencion")) = "si" Then
    Me.ChkRetencion.Value = 1
    StrRetencion = "si"
   Else
    Me.ChkRetencion.Value = 0
    StrRetencion = "no"
  End If
  
  If rst("cliente_mayor") = "si" Then
    Me.chkclientemayor.Value = 1
  Else
    Me.chkclientemayor.Value = 0
  End If
  
  If Trim(rst("id_auspeciador")) = "si" Then
    Me.ChkAuspiciador.Value = 1
    StrAuspiciador = "si"
   Else
    Me.ChkAuspiciador.Value = 0
    StrAuspiciador = "no"
  End If
  If Trim(rst("id_almacen")) = "si" Then
    Me.ChkAlmacen.Value = 1
    StrAlmacen = "si"
   Else
    Me.ChkAlmacen.Value = 0
    StrAlmacen = "no"
  End If
  
  If Trim(rst("id_cliente")) = "si" Then
    Me.ChkCliente.Value = 1
    StrCliente = "si"
   Else
    Me.ChkCliente.Value = 0
    StrCliente = "no"
  End If
  If Trim(rst("id_proveedor")) = "si" Then
    Me.chkProveedor.Value = 1
    strProveedor = "si"
   Else
    Me.chkProveedor.Value = 0
    strProveedor = "no"
  End If
  '----------------------------------------
  If Trim(rst("id_personal")) = "no" Then
   
   
    Me.ChkPersonal.Value = 0
    StrPersonal = "no"
    ' Me.SstKardex.TabVisible(1) = False
  End If
  '-------------------------------------------
  Me.TxtLicencia.Text = rst("licencia")
  If Trim(rst("id_transporte")) = "si" Then
    Me.ChkTransporte.Value = 1
    
    
    
    StrTransporte = "si"
   Else
    Me.ChkTransporte.Value = 0
    StrTransporte = "no"
  End If
  If IsNull(rst("mail")) = False Then
      Me.txtEmail.Text = rst("mail")
    Else
    Me.txtEmail.Text = ""
  End If
  
If (rst("id_departamento") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lbldepartamento(1).Visible = False
    Me.DtcDepartamento.Visible = False
End If

If (rst("id_provincia") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lblprovincia(0).Visible = False
    Me.DtcProvincia.Visible = False
End If

If Trim(rst("id_credito")) = "si" Then
    If Len((rst("id_empresa_credito"))) = 11 Then
        Me.chkEmpresa.Value = True
        Me.txtRucEmpresa.Text = rst("id_empresa")
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRucEmpresa.Text) & "' LIMIT 1"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            Me.LblEmpresa.Caption = UCase(rstTemporal("nombre_completo"))
        End If
     Else
      
        Me.ChkMaximoCredito.Value = True
        Me.txtMaximoCredito.Text = Format(rst("monto_credito"), "###0.00")
    End If
Else

    Me.OptSincredito.Value = True
End If

'--------- foto--------
If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
    If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
        Me.Image1.Visible = True
        Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
        img = Trim(rst("foto"))
    Else
        Me.Image1 = Nothing
    End If
End If
'--------- foto--------

'If Val(cDepartamento) > 0 Then
'    strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & cDepartamento & "'"
'    Call ConfiguraRst(strCadena)
'    Call LlenaDataCombo(Me.DtcDepartamento)
'    Me.DtcDepartamento.BoundText = cDepartamento
'End If

'If Val(cProvincia) > 0 Then
'    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE  id_departamento='" & cDepartamento & "'"
'    Call ConfiguraRst(strCadena)
'    Call LlenaDataCombo(Me.DtcProvincia)
'    Me.DtcProvincia.BoundText = cProvincia
'End If

'If Val(cDistrito) > 0 Then
'    strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE  id_provincia='" & cProvincia & "'"
'    Call ConfiguraRst(strCadena)
'    Call LlenaDataCombo(Me.DtcDistrito)
'    Me.DtcDistrito.BoundText = cDistrito
'End If

'Call CargarLogo(StrCodTabla)
Call LlenarTelefonos(Me.HfTelefonos, StrCodTabla)
Call Me.llenar_direccion(Me.hfdireccion, Trim(Me.txtRuc.Text))

If KEY_RUBRO = "00025" Then
    Call load_estudiante(Trim(Me.txtRuc.Text), Val(Me.lbl_id_matricula.Caption))
End If

'Call load_pais(in_pais)
Call load_seguro_persona(in_seguro, Trim(Me.txtRuc.Text))

Call get_zona(in_zona)
Call Me.get_ubigeo_sunat(codigo_ubigeo_sunat)


Exit Function
'salir:   MsgBox "Se Presento un Problema Disculpe las molestias", vbInformation, KEY_EMPRESA
End Function
Private Sub load_seguro_lista()
strCadena = "SELECT id_detalle as Codigo,Descripcion as Descripcion FROM seguro_medico_detalle WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.Dtcseguroempresa)
End Sub
Private Sub load_seguro_persona(ByVal in_seguro As String, ByVal in_dni As String)

If in_seguro = "si" Then
    Me.chk_seguro_transporte.Value = 1
    'Call load_seguro_lista
    strCadena = "SELECT * FROM persona_seguro WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.Dtcseguroempresa.BoundText = rst("id_detalle")
       Me.txtpolizaseguro_transporte.Text = rst("numero")
       Me.DtpFechaEmision_transporte.Value = rst("expedicion")
       Me.DtpFechaCaducidad_transporte.Value = rst("expiracion")
       If rst("activo") = "si" Then
          Me.chk_estadoseguro.Value = 1
       Else
          Me.chk_estadoseguro.Value = 0
       End If
    End If
End If
End Sub
Private Sub load_pais(ByVal in_pais As String)
  If in_pais <> "pe" Then
  strCadena = "SELECT id_pais as Codigo,descripcion as Descripcion FROM pais WHERE id_pais='pe' LIMIT 1"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPais)
  End If
End Sub
Private Sub llenarFamiliares(ByVal Grilla As MSHFlexGrid)
strCadena = "SELECT * FROM view_familiar WHERE dni='" & Me.txtRuc.Text & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    
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
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 3500
        Next
        cabecera = "IDCODIGO" & vbTab & "DNI" & vbTab & "NOMBRE COMPLETO" & vbTab & "PARENTESCO" & vbTab & "OCUPACION" & vbTab & "GRADO INSTRUCCION" & vbTab & "TELEFONO" & vbTab & "DIRECCION"
        Grilla.AddItem cabecera
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id") & vbTab & rstT("dni_familia") & vbTab & rstT("nombre_completo") & vbTab & rstT("parentesco") & vbTab & rstT("ocupacion") & vbTab & rstT("grado") & vbTab & rstT("telefono") & vbTab & rstT("direccion")
          Grilla.AddItem Fila
          
          rstT.MoveNext
      Next i
   
End Sub

Public Sub LLENA_NC(ByVal cPersona As String)
'On Error GoTo salir
Dim cDepartamento As String, cProvincia As String, cDistrito As String, cUrbanizacion As Double, cZona As Double
strCadena = "SELECT * FROM persona P WHERE  P.dni = '" & cPersona & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Exit Sub
End If
   StrCodTabla = rst("dni")
  Me.LblCodPersona.Caption = StrCodTabla
  If IsNull(rst("dni")) = True Then
    GoTo sin_doc
  End If
  Me.txtPaterno.Text = UCase(rst("a_paterno"))
  Me.txtMaterno.Text = UCase(rst("a_materno"))
  Me.txtNombre.Text = UCase(rst("nombres"))
  
  Me.TxtLicencia.Text = rst("licencia")
  If Len(rst("dni")) = 8 Then
         Me.LblEntidad.Caption = "Nombre:"
          Me.LblTipoDocumento.Caption = "DNI"
    Else
        Me.LblEntidad.Caption = "Razon Social:"
        Me.LblTipoDocumento.Caption = "RUC"
    End If
sin_doc:
    Me.TxtRazonSocial.Text = rst("nombre_completo")
        
If (rst("descuento") > 0) Then
    Me.chkDescuento.Value = 1
    Me.txtDescuento.Visible = True
    Me.txtDescuento.Text = rst("descuento")
Else
    Me.chkDescuento.Value = 0
End If

    
    Me.txtdia.Visible = True
    Me.DtcMes.Visible = True
    Me.txtdia.Text = formato_item(rst("id_dia"), 2)
    Me.DtcMes.BoundText = formato_item(rst("id_mes"), 2)
    cDepartamento = rst("id_departamento")
    cProvincia = rst("id_provincia")
    cDistrito = rst("id_distrito")
    'cUrbanizacion = rst("id_urbanizacion")
    'cZona = rst("id_zona")
    Me.TxtDireccion1.Text = rst("direccion")
    Me.txtRuc.Text = rst("dni")
  
  If IsNull(rst("mail")) = False Then
      Me.txtEmail.Text = rst("mail")
    Else
    Me.txtEmail.Text = ""
  End If
  
If (rst("id_departamento") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lbldepartamento(0).Visible = False
    Me.DtcDepartamento.Visible = False
End If

If (rst("id_provincia") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lblprovincia(0).Visible = False
    Me.DtcProvincia.Visible = False
End If


If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
    'Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + Trim(rst("foto")))
    img = Trim(rst("foto"))
End If
'--------- foto--------


If Val(cDepartamento) > 0 Then
    strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & cDepartamento & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDepartamento)
    Me.DtcDepartamento.BoundText = cDepartamento
    Set rst = Nothing
End If

If Val(cProvincia) > 0 Then
    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_provincia='" & cProvincia & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProvincia)
    Me.DtcProvincia.BoundText = cProvincia
    Set rst = Nothing
End If



'Call CargarLogo(StrCodTabla)
strCadena = "SELECT * FROM  persona_telefono T,persona_cargos C WHERE T.dni='" & StrCodTabla & "' AND T.id_cargo=C.id_cargo"
Call LlenarTelefonos(Me.HfTelefonos, StrCodTabla)

  Exit Sub
'salir:   MsgBox "Se Presento un Problema Disculpe las molestias", vbInformation, KEY_EMPRESA
End Sub
Private Sub CargarLogo(ByVal cPersona As String)
Dim sql As String
Dim sw As String
sql = "select foto From persona Where dni='" & Trim(cPersona) & "'"
Call ConfiguraRst(sql)
If rst.RecordCount > 0 Then

If IsNull(rst(0)) = False Then
    Image1.Picture = Leer_Imagen(CnBd, sql, "foto")
End If
End If
Set rst = Nothing
End Sub

Private Sub HfChofer_SelChange()
If Me.HfChofer.Rows > 0 Then
   Me.cmdEliminarchofer.Enabled = True
Else
   Me.cmdEliminarchofer.Enabled = False
End If
End Sub

Private Sub HfDireccion_SelChange()
If Val(Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0)) > 0 Then
   Me.cmdmodificar_direccion.Enabled = True
   Me.cmdeliminar_direccion.Enabled = True
Else
    Me.cmdmodificar_direccion.Enabled = False
   Me.cmdeliminar_direccion.Enabled = False
End If
End Sub





Private Sub HfMarcas_SelChange()
If Val(Me.HfMarcas.TextMatrix(Me.HfMarcas.Row, 0)) > 0 Then
   Me.cmdEliminaMarca.Enabled = True
Else
   Me.cmdEliminaMarca.Enabled = False
End If
End Sub

Private Sub Hfplanservicio_SelChange()
If Me.Hfplanservicio.Rows > 0 Then
    If Val(Me.Hfplanservicio.TextMatrix(Me.Hfplanservicio.Row, 0)) > 0 Then
        Me.cmdeliminarplan.Enabled = True
        Me.cmdmodificarplan.Enabled = True
    Else
        Me.cmdeliminarplan.Enabled = False
        Me.cmdmodificarplan.Enabled = False
    End If
End If
End Sub

Private Sub HfTelefonos_SelChange()
If Me.HfTelefonos.Rows > 0 Then
    If Trim(Me.HfTelefonos.TextMatrix(Me.HfTelefonos.Row, 2)) <> "TELEFONO" Then
        Me.cmdeditarTelefono.Enabled = True
    Else
        Me.cmdeditarTelefono.Enabled = False
    End If
Else
    Me.cmdeditarTelefono.Enabled = False
End If
End Sub

Private Sub HfUbigeo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.HfUbigeo.Rows > 0 Then
       
       If Procedencia = buscar Then
            Me.lblCodigoUbigeoSunat.Caption = Me.HfUbigeo.TextMatrix(Me.HfUbigeo.Row, 0)
            Call get_ubigeo_sunat(Me.lblCodigoUbigeoSunat.Caption)
       Else
            Me.lblCodigoUbigeoSunat2.Caption = Me.HfUbigeo.TextMatrix(Me.HfUbigeo.Row, 0)
            Call get_ubigeo_sunat2(Me.lblCodigoUbigeoSunat2.Caption)
       End If
       
       
       
        Me.frmUbigeo.Visible = False
    End If
End If
End Sub

Public Sub get_ubigeo_sunat2(ByVal in_codigo As String)

strCadena = "SELECT cod_dep_sunat as Codigo,desc_dep_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcDepartamento2)

strCadena = "SELECT cod_prov_sunat as Codigo,desc_prov_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcProvincia2)

strCadena = "SELECT cod_ubigeo_sunat as Codigo,desc_ubigeo_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcDistrito2)


End Sub
Public Sub get_ubigeo_sunat(ByVal in_codigo As String)

strCadena = "SELECT cod_dep_sunat as Codigo,desc_dep_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcDepartamento)

strCadena = "SELECT cod_prov_sunat as Codigo,desc_prov_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcProvincia)

strCadena = "SELECT cod_ubigeo_sunat as Codigo,desc_ubigeo_sunat as Descripcion FROM ubigeo WHERE cod_ubigeo_sunat='" & in_codigo & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcDistrito)


End Sub
Private Sub Image2_Click()
Me.frmUbigeo.Visible = False
End Sub

Private Sub OptSincredito_Click()
If Me.OptSincredito.Value = True Then
    Me.LblEmpresa.Visible = False
    Me.txtRucEmpresa.Visible = False
    Me.txtMaximoCredito.Visible = False
    
    
    
End If
End Sub




Private Sub SstKardex_Click(PreviousTab As Integer)
If Me.SstKardex.Tab = 1 Then
    Call Me.llenar_plan_servicio(Me.Hfplanservicio, Trim(Me.txtRuc.Text))
End If

If Me.SstKardex.Tab = 5 Then
     Call llenarFamiliares(Me.HfgFamiliares)
End If
End Sub



Private Sub TxtNDocumento_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtEmail)
End If
End Sub

Private Sub txtBuscaUbigeo_Change()
If Len(Trim(Me.txtBuscaUbigeo.Text)) >= 3 Then
    Call llenar_ubigeo(Me.HfUbigeo, Trim(Me.txtBuscaUbigeo.Text))
End If
End Sub

Private Sub txtchofer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
buscar_nuevamente:
    strCadena = "SELECT nombre_completo, licencia FROM persona WHERE dni='" & Trim(Me.txtchofer.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.lblchofer.Text = rst("nombre_completo")
       Me.txtlicenciaTransporte.Text = rst("licencia")
    Else
    
        If Len(Trim(Me.txtchofer.Text)) = 8 Then
            If get_dni_reniec_ii(Trim(Me.txtchofer.Text)) = True Then
                GoTo buscar_nuevamente
            End If
        End If
        
       
       Me.lblchofer.Text = ""
       Me.txtlicenciaTransporte.Text = ""
    End If
    Exit Sub
End If
End Sub

Private Sub txtcodigo_plan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmProducto.Show
   Exit Sub
End If
End Sub

Private Sub txtdireccion3_Change()
'If Trim(Me.txtdireccion3.Text) <> "" Then
'    strCadena = "SELECT id_distrito as Codigo,CONCAT(d.descripcion,' - ',p.descripcion) as Descripcion FROM distrito d,provincia p  WHERE  d.id_provincia=p.id_provincia and  d.descripcion LIKE '%" & Trim(Me.txtdireccion3.Text) & "%'"
'    Call ConfiguraRst(strCadena)
'    Call LlenaDataCombo(Me.DtcDistrito2)
'    Set rst = Nothing
'End If
End Sub

Private Sub txtdireccion3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  If Me.DtcDistrito2.BoundText <> "" Then
  '      Me.DtcDistrito2.SetFocus
  '  End If
  Procedencia = Selecionar
  

  
  Me.frmUbigeo.Visible = True
    Me.frmUbigeo.Top = Me.txtdireccion3.Top
    Call Resalta(Me.txtBuscaUbigeo)
  
  
End If
End Sub


Private Sub txtsectorista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtsectorista.Text) & "%' and  ruc='" & KEY_RUC & "' AND id_personal='si' ORDER BY nombre_completo LIMIT 1"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcVendedor)
End If
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Sub verificaTipo()
If Me.ChkCliente.Value = 1 Then
    StrCliente = "si"
Else
    StrCliente = "no"
End If
If Me.ChkAuspiciador.Value = 1 Then
    StrAuspiciador = "si"
Else
    StrAuspiciador = "no"
End If

If (Me.chkDescuento.Value = 1) Then
    descuento_por = Me.txtDescuento.Text
Else
    descuento_por = 0
End If

If Me.chkProveedor.Value = 1 Then
    strProveedor = "si"
Else
    strProveedor = "no"
End If
If Me.ChkContable.Value = 1 Then
    StrContable = "si"
Else
    StrContable = "no"
End If
If Me.ChkTransporte.Value = 1 Then
    StrTransporte = "si"
Else
    StrTransporte = "no"
End If
If Me.ChkPersonal.Value = 1 Then
    StrPersonal = "si"
Else
    StrPersonal = "no"
End If
If Me.ChkAlmacen.Value = 1 Then
    StrAlmacen = "si"
Else
    StrAlmacen = "no"
End If

If Me.ChkPercepcion.Value = 1 Then
    StrPercepcion = "si"
Else
    StrPercepcion = "no"
End If
If Me.ChkRetencion.Value = 1 Then
    StrRetencion = "si"
Else
    StrRetencion = "no"
End If
End Sub




Private Sub txtCelular_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdia)
End If
End Sub

Private Sub txtdia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMes.SetFocus
End If
End Sub

Private Sub TxtDireccion1_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdia)
End If
End Sub

Private Sub TxtDireccion2_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdia)
End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtFmaterno.Text = rst("a_materno")
       Me.TxtFpaterno.Text = rst("a_paterno")
       Me.TxtFnombers.Text = rst("nombres")
       Me.dtcparentesco.SetFocus
    Else
        Call Resalta(Me.TxtFpaterno)
    End If
    
End If
End Sub

Private Sub TxtDistrito_Change()
'If Trim(Me.TxtDistrito.Text) <> "" Then
'
'    strCadena = "SELECT id_distrito as Codigo,CONCAT(d.descripcion,' - ',p.descripcion) as Descripcion FROM distrito d,provincia p  WHERE  d.id_provincia=p.id_provincia and  d.descripcion LIKE '%" & Trim(Me.TxtDistrito.Text) & "%'"
'    Call ConfiguraRst(strCadena)
'    Call LlenaDataCombo(Me.DtcDistrito)
'    Set rst = Nothing
'
'End If
End Sub

Private Sub TxtDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    Me.frmUbigeo.Visible = True
    Me.frmUbigeo.Top = Me.TxtDistrito.Top - Val(Me.frmUbigeo.Height) + Val(Me.TxtDistrito.Height)
    Call Resalta(Me.txtBuscaUbigeo)
    
    
    'If Me.DtcDistrito.BoundText <> "" Then
    '    Me.DtcDistrito.SetFocus
    'End If
End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtFono)
End If
End Sub

Private Sub TxtEntidad_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtDireccion1.SetFocus
End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtRuc.SetFocus
End If
End Sub


Private Sub TxtFono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtFono.Text <> "" And Len(Me.TxtFono.Text) > 0 Then
        Me.Command1.SetFocus
    Else
        If MsgBox("No desea Ingresar el Telefono de Contacto", vbYesNo + vbQuestion, "Mensaje para el Usuario") = vbYes Then
            Call Resalta(Me.TxtFono)
        Else
            Call Resalta(Me.TxtDistrito)
        End If
    End If
End If
End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDistrito)
End If
End Sub

Private Sub txtMaterno_Change()
Me.TxtRazonSocial.Text = ""
Me.TxtRazonSocial.Text = UCase(Trim(Me.txtPaterno.Text)) + Space(1) + UCase(Trim(Me.txtMaterno.Text)) + Space(1) + UCase(Trim(Me.txtNombre.Text))
End Sub

Private Sub txtMaterno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtNombre)
End If
End Sub

Private Sub txtNombre_Change()
Me.TxtRazonSocial.Text = ""
Me.TxtRazonSocial.Text = UCase(Trim(Me.txtPaterno.Text)) + Space(1) + UCase(Trim(Me.txtMaterno.Text)) + Space(1) + UCase(Trim(Me.txtNombre.Text))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDireccion1)
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtFono)
End If
End Sub

Private Sub txtPaterno_Change()
Me.TxtRazonSocial.Text = ""
Me.TxtRazonSocial.Text = UCase(Trim(Me.txtPaterno.Text)) + Space(1) + UCase(Trim(Me.txtMaterno.Text)) + Space(1) + UCase(Trim(Me.txtNombre.Text))
End Sub

Private Sub txtPaterno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtMaterno)
End If
End Sub

Private Sub TxtRuc_Change()
If FrmPersona.Procedencia <> modificar Then
    strCadena = "SELECT dni FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Me.div_verifica.Visible = True
        Set rstT = Nothing
        strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Me.lblresultado.Caption = "Entidad ya forma parte de sus Clientes"
            Me.cmdVisualizar.Visible = True
        Else
            Me.lblresultado.Caption = "Registrado en www.Vitekey.com" + Chr(13)
            Me.cmdVisualizar.Visible = True
        End If
    Else
        Me.div_verifica.Visible = False
    End If
    Set rstT = Nothing
    End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call Resalta(Me.txtdia)
 
    Me.SstKardex.TabVisible(5) = True

 
 
 strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Call LLENA(rst("dni"))
    Call precionar
 Else
    Call precionar
 End If
 

End If
End Sub

Private Sub txtRucEmpresa_Change()
If Len(Me.txtRucEmpresa.Text) = 11 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRucEmpresa.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.LblEmpresa.Caption = UCase(rst("nombre_completo"))
        Me.LblEmpresa.Visible = True
    Else
        MsgBox "Empresa no Registrada Fabor de Registrar", vbInformation, "Mensaje para el Usuario"
        Me.LblEmpresa.Caption = ""
    End If
    Set rst = Nothing
Else
    Me.LblEmpresa.Caption = ""
End If
End Sub

Private Sub txtRucEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtRucEmpresa.Text = "" Then
        Procedencia = buscar
        FrmPersona.Show
    End If
End If
End Sub

Private Sub TxtTelefono1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtTelefono2.SetFocus
End If
End Sub

Private Sub TxtTelefono2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtFax.SetFocus
End If
End Sub

Private Sub txtZona_Change()
    strCadena = "SELECT id_zona as Codigo,descripcion_zona as Descripcion  FROM zona WHERE descripcion_zona LIKE '%" & Trim(Me.txtZona.Text) & "%'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcZona)
End Sub

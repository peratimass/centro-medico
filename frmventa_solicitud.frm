VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmventa_solicitud 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "COMPRADOR"
      TabPicture(0)   =   "frmventa_solicitud.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblconyuge"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDni_conyuge"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtvivienda"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtmonto_alquiler"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProfesion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtIngreso_mensual"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtIngresos_conyuge"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOtros_ingresos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtTotal_ingresos"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "FIADOR [GARANTE]"
      TabPicture(1)   =   "frmventa_solicitud.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblfiafor"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txttelefono_fiador"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtruc_fiador"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtingresos_fiador"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtprofesion_fiador"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtOtrosIngresos_fiador"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtTotal_ingreso_fiador"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtTotal_ingreso_fiador 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   1920
         TabIndex        =   31
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtOtrosIngresos_fiador 
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
         Left            =   4440
         TabIndex        =   29
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtprofesion_fiador 
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
         Left            =   1920
         TabIndex        =   26
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtingresos_fiador 
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
         Left            =   1920
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtruc_fiador 
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
         Left            =   1920
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txttelefono_fiador 
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
         Left            =   1920
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtTotal_ingresos 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         TabIndex        =   18
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtOtros_ingresos 
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
         TabIndex        =   16
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtIngresos_conyuge 
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
         Left            =   -72120
         TabIndex        =   14
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtIngreso_mensual 
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
         Left            =   -72120
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtProfesion 
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
         Left            =   -73560
         TabIndex        =   11
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtmonto_alquiler 
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
         Left            =   -68880
         TabIndex        =   10
         Top             =   1605
         Width           =   1575
      End
      Begin VB.TextBox txtvivienda 
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
         Left            =   -73560
         TabIndex        =   8
         Top             =   1605
         Width           =   2535
      End
      Begin VB.TextBox txtDni_conyuge 
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
         Left            =   -73560
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL INGRESOS :"
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
         Left            =   600
         TabIndex        =   32
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OTROS ING :"
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
         Left            =   3240
         TabIndex        =   30
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROFESION :"
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
         Left            =   480
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ING.MENSUAL :"
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
         Left            =   600
         TabIndex        =   27
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblfiafor 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   23
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO :"
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
         Left            =   600
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL INGRESOS :"
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
         Left            =   -70800
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OTROS ING :"
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
         Left            =   -70680
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ING CONYUGE :"
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
         Left            =   -73440
         TabIndex        =   15
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ING.MENSUAL :"
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
         Left            =   -73440
         TabIndex        =   12
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO ALQUILER:"
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
         Left            =   -70560
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   -73440
         X2              =   -66840
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblconyuge 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72240
         TabIndex        =   7
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CONYUGE :"
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
         Left            =   -74640
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PROFESION :"
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
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VIVIENDA :"
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
         Left            =   -74640
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   3375
         Left            =   240
         Top             =   480
         Width           =   8055
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   3495
         Left            =   -74880
         Top             =   360
         Width           =   8295
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1296
      btype           =   5
      tx              =   "PROCESAR"
      enab            =   -1  'True
      font            =   "frmventa_solicitud.frx":0038
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmventa_solicitud.frx":0060
      picn            =   "frmventa_solicitud.frx":007E
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   735
      Left            =   7560
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1296
      btype           =   5
      tx              =   "CERRAR"
      enab            =   -1  'True
      font            =   "frmventa_solicitud.frx":36C6
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmventa_solicitud.frx":36EE
      picn            =   "frmventa_solicitud.frx":370C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   735
      Left            =   6240
      TabIndex        =   34
      Top             =   4320
      Width           =   1215
      _extentx        =   2143
      _extenty        =   1296
      btype           =   5
      tx              =   "IMPRIMIR"
      enab            =   -1  'True
      font            =   "frmventa_solicitud.frx":6724
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmventa_solicitud.frx":674C
      picn            =   "frmventa_solicitud.frx":676A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label lblsolicitud 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   480
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000015&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   5205
      Left            =   0
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmventa_solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCerrar_Click()
    Call enabled_form(FrmVentas)
    Unload Me
    Exit Sub
End Sub


Private Sub cmdImprimir_Click()
On Error GoTo salir
strCadena = "SELECT * FROM view_solicitud_credito_ii WHERE id_venta='" & Val(FrmVentas.TxtIdVenta.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_solicitud_credito", , App.Path + "\Reportes\")
Exit Sub
salir:
End Sub

Private Sub cmdProcesar_Click()

If Val(Me.txtIngreso_mensual.Text) > 0 Then
  
  strCadena = "call sp_solicitud_venta('" & Val(Me.lblsolicitud.Caption) & "','" & Val(FrmVentas.TxtIdVenta.Text) & "','" & FrmVentas.TxtCodCliente.Text & "'" & _
  ",'" & Trim(Me.txtDni_conyuge.Text) & "','" & Trim(Me.txtruc_fiador.Text) & "','" & Trim(Me.txtvivienda.Text) & "','" & Val(Me.txtmonto_alquiler.Text) & "' " & _
  ",'" & Val(Me.txtIngreso_mensual.Text) & "','" & Val(Me.txtIngresos_conyuge.Text) & "','" & Val(Me.txtOtros_ingresos.Text) & "','-','0','" & Trim(Me.txtprofesion_fiador.Text) & "' " & _
  ",'" & Trim(Me.txtProfesion.Text) & "','" & Val(Me.txtingresos_fiador.Text) & "','" & Val(Me.txtOtrosIngresos_fiador.Text) & "')"
  CnBd.Execute (strCadena)
  
End If


End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100


End Sub

Public Sub put_llenar(ByVal in_venta As String)

INICIO:
strCadena = "SELECT * FROM view_solicitud_credito_ii WHERE id_venta='" & Val(in_venta) & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.lblsolicitud.Caption = rst("id_solicitud")
   Me.txtDni_conyuge.Text = rst("dni_conyuge")
   If IsNull(rst("nombre_completo")) = True Then
        Me.lblconyuge.Caption = "-"
   Else
        Me.lblconyuge.Caption = rst("nombre_completo")
   End If
   Me.txtvivienda.Text = rst("vivienda")
   Me.txtmonto_alquiler.Text = rst("alquiler")
   Me.txtProfesion.Text = rst("profesion_comprador")
   Me.txtIngreso_mensual.Text = rst("ingreso_comprador")
   Me.txtIngresos_conyuge.Text = rst("ingreso_conyuge")
   Me.txtOtros_ingresos.Text = rst("otro_ingreso_comprador")
   Me.txtruc_fiador.Text = rst("dni_fiador")
   If IsNull(rst("fiador")) = True Then
        Me.lblfiafor.Caption = "-"
   Else
        Me.lblfiafor.Caption = rst("fiador")
   End If
   
   If IsNull(rst("celular_fiador")) = True Then
       Me.txttelefono_fiador.Text = " "
   Else
        Me.txttelefono_fiador.Text = rst("celular_fiador")
   End If
   
   Me.txtprofesion_fiador.Text = rst("profesion_fiador")
   Me.txtingresos_fiador.Text = rst("ingreso_fiador")
   Me.txtOtrosIngresos_fiador.Text = rst("otros_ingreso_fiador")
   
Else
    strCadena = "call sp_solicitud_venta('0','" & Val(in_venta) & "','" & FrmVentas.TxtCodCliente.Text & "','0','0','-','0','0','0','0','-','0','-','-','0','0')"
    CnBd.Execute (strCadena)
    GoTo INICIO
End If



End Sub

Private Sub txtDni_conyuge_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.lblconyuge.Caption = get_persona(Trim(Me.txtDni_conyuge.Text))
End If
End Sub

Private Sub txtIngreso_mensual_Change()
Me.txtTotal_ingresos.Text = Format(Val(Me.txtIngreso_mensual.Text) + Val(Me.txtIngresos_conyuge.Text) + Val(Me.txtOtros_ingresos.Text), "###0.00")
End Sub

Private Sub txtIngresos_conyuge_Change()
Me.txtTotal_ingresos.Text = Format(Val(Me.txtIngreso_mensual.Text) + Val(Me.txtIngresos_conyuge.Text) + Val(Me.txtOtros_ingresos.Text), "###0.00")
End Sub

Private Sub txtingresos_fiador_Change()
Me.txtTotal_ingreso_fiador.Text = Format(Val(Me.txtingresos_fiador.Text) + Val(Me.txtOtrosIngresos_fiador.Text), "###0.00")
End Sub

Private Sub txtOtros_ingresos_Change()
Me.txtTotal_ingresos.Text = Format(Val(Me.txtIngreso_mensual.Text) + Val(Me.txtIngresos_conyuge.Text) + Val(Me.txtOtros_ingresos.Text), "###0.00")
End Sub

Private Sub txtruc_fiador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.lblfiafor.Caption = get_persona(Trim(Me.txtruc_fiador.Text))
End If
End Sub

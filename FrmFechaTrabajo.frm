VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmFechaTrabajo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12345
   Icon            =   "FrmFechaTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DtcEmpresa 
      Height          =   405
      Left            =   4320
      TabIndex        =   27
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frmactivacion 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1725
      Left            =   4365
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   7815
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
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   2950
         TabIndex        =   25
         Top             =   1995
         Width           =   3000
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2160
         Top             =   720
      End
      Begin VitekeySoft.ChameleonBtn CmdActivacion 
         Height          =   1530
         Left            =   1920
         TabIndex        =   23
         Top             =   120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2699
         BTYPE           =   5
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmFechaTrabajo.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdactivar 
         Height          =   405
         Left            =   6000
         TabIndex        =   14
         Top             =   1965
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "ACTIVAR"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmFechaTrabajo.frx":0028
         PICN            =   "FrmFechaTrabajo.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACTIVACION:"
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
         Height          =   195
         Left            =   2040
         TabIndex        =   24
         Top             =   2040
         Width           =   870
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   735
         Left            =   1920
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Image camara1 
         Height          =   1530
         Left            =   120
         Picture         =   "FrmFechaTrabajo.frx":0496
         Top             =   120
         Width           =   1785
      End
      Begin VB.Image camara2 
         Height          =   1530
         Left            =   120
         Picture         =   "FrmFechaTrabajo.frx":553D
         Top             =   120
         Width           =   1785
      End
   End
   Begin VB.Frame frmprincipal 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      Begin VB.TextBox txtEmpresa 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   10560
         TabIndex        =   28
         Top             =   3840
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox chkImpresora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "PRINTER'S"
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
         Height          =   195
         Left            =   6120
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkFinger 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "FINGERS"
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
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox chkCamara 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CAMARA"
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
         Height          =   195
         Left            =   3960
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   7680
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox TxtImpresora 
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
            Left            =   2040
            TabIndex        =   4
            Text            =   "\\"
            Top             =   240
            Width           =   2415
         End
         Begin VitekeySoft.ChameleonBtn cmdImpresora 
            Height          =   375
            Left            =   2040
            TabIndex        =   5
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "AGREGAR IMPRESORA"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmFechaTrabajo.frx":A683
            PICN            =   "FrmFechaTrabajo.frx":A69F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOMBRE IMPRESORA"
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
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1650
         End
      End
      Begin VB.CheckBox chkinvitado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "INVITADO"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   250
         Left            =   10500
         TabIndex        =   2
         Top             =   5355
         Width           =   1215
      End
      Begin VB.CommandButton cImpresora 
         Height          =   735
         Left            =   6360
         Picture         =   "FrmFechaTrabajo.frx":A72C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   405
         Left            =   4320
         TabIndex        =   10
         Top             =   4560
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MMM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   390
         Left            =   4320
         TabIndex        =   11
         Top             =   6480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   688
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   186646529
         CurrentDate     =   40692
      End
      Begin MSDataListLib.DataCombo DtcTurno 
         Height          =   405
         Left            =   4320
         TabIndex        =   12
         Top             =   5835
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   714
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminarImpresora 
         Height          =   315
         Left            =   11880
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   315
         _extentx        =   556
         _extenty        =   556
         btype           =   5
         tx              =   ""
         enab            =   -1
         font            =   "FrmFechaTrabajo.frx":DF7A
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   8388608
         fcolo           =   8388608
         mcol            =   12632256
         mptr            =   1
         micon           =   "FrmFechaTrabajo.frx":DFA2
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   1
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin MSDataListLib.DataCombo DtcVentanilla 
         Height          =   405
         Left            =   4320
         TabIndex        =   15
         Top             =   5280
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn Command2 
         Height          =   450
         Left            =   8840
         TabIndex        =   16
         Top             =   6435
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         btype           =   3
         tx              =   "ACEPTAR "
         enab            =   -1
         font            =   "FrmFechaTrabajo.frx":DFC0
         coltype         =   2
         focusr          =   -1
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "FrmFechaTrabajo.frx":DFE8
         picn            =   "FrmFechaTrabajo.frx":E006
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VitekeySoft.ChameleonBtn Command1 
         Height          =   450
         Left            =   7125
         TabIndex        =   17
         Top             =   6435
         Width           =   1575
         _extentx        =   2990
         _extenty        =   873
         btype           =   3
         tx              =   "ACTUALIZAR"
         enab            =   -1
         font            =   "FrmFechaTrabajo.frx":1054A
         coltype         =   1
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   0
         fcolo           =   0
         mcol            =   12632256
         mptr            =   1
         micon           =   "FrmFechaTrabajo.frx":10572
         picn            =   "FrmFechaTrabajo.frx":10590
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPrinter 
         Height          =   1695
         Left            =   7680
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2990
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
      Begin VB.Image imgPais 
         Height          =   810
         Left            =   10560
         Picture         =   "FrmFechaTrabajo.frx":1379C
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   6375
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Height          =   7020
         Left            =   0
         Top             =   0
         Width           =   12345
      End
      Begin VB.Label lblHoras 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   7080
         TabIndex        =   21
         Top             =   5835
         Width           =   3345
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4020
         TabIndex        =   20
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblDni 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4020
         TabIndex        =   19
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label lblcargo 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4020
         TabIndex        =   18
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   3375
         Left            =   240
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   5160
         Picture         =   "FrmFechaTrabajo.frx":16DE2
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   5160
         Picture         =   "FrmFechaTrabajo.frx":1A0B6
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   495
      End
      Begin VB.Image Image4 
         Height          =   7680
         Left            =   0
         Picture         =   "FrmFechaTrabajo.frx":1D3AA
         Top             =   0
         Width           =   13185
      End
   End
End
Attribute VB_Name = "FrmFechaTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procendencia As EnumProcede

Private Sub chkImpresora_Click()
If Me.chkImpresora.Value = 1 Then
    Me.cImpresora.Visible = True
    Me.HfPrinter.Visible = True
    Call Me.llenar_printer(Me.HfPrinter, Me.DtcEmpresa.BoundText)
Else
    Me.cImpresora.Visible = False
    
    Me.HfPrinter.Visible = False
    Me.cmdEliminarImpresora.Visible = False
End If
End Sub
Public Function verificacion_activacion() As Boolean
strCadena = "SELECT  CURDATE()"
Call ConfiguraRstK(strCadena)

ultimafecha = Format(rstK(0), "YYYY-mm-dd")

strCadena = "SELECT * FROM entidad_parametros WHERE cod_unico='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst("activacion_permanente") = "no" Then
   If IsNull(rst("caducidad")) = False Then
        dias = DateDiff("d", ultimafecha, Format(rst("caducidad"), "YYYY-mm-dd"))
        If dias >= 0 Then
            Me.CmdActivacion.Visible = True
            Me.CmdActivacion.Caption = "REGULARICE SUS PAGOS DEL SOFTWARE" + Space(1) + "LE QUEDAN" + Space(2) + str(dias) + Space(2) + " DIAS DE USO DEL SISTEMA"
            Me.frmprincipal.Enabled = True
            Me.frmactivacion.Height = 1725
            Me.frmactivacion.Visible = True
            Exit Function
        Else
            Me.frmprincipal.Enabled = False
            Me.CmdActivacion.Caption = "REGULARICE SUS PAGOS DEL SOFTWARE" + Chr(13) + "CELL:942867953"
            Me.frmactivacion.Visible = True
            MDIFrmPrincipal.Toolbar1.Enabled = False
            MDIFrmPrincipal.MnuInformesContables.Enabled = False
            MDIFrmPrincipal.MnuInformesContables.Enabled = False
            MDIFrmPrincipal.mnucaja.Enabled = False
            MDIFrmPrincipal.mnuInformessunat.Enabled = False
            MDIFrmPrincipal.mnuInformesGerenciales = False
            MDIFrmPrincipal.MnuGestionFinanciera = False
            MDIFrmPrincipal.mnuGestionegocio = False
        End If
    Else
        Me.CmdActivacion.Visible = False
   End If
   Else
   Me.CmdActivacion.Visible = False
End If
End Function
Private Sub cImpresora_Click()
If Me.Frame2.Visible = True Then
    Me.Frame2.Visible = False
    'Me.HfPrinter.Visible = False
Else
    Me.Frame2.Visible = True
    Me.HfPrinter.Visible = True
End If
End Sub

Private Sub CmdActivacion_Click()
Me.frmactivacion.Height = 2565
Call Resalta(Me.TxtPassword)
Call get_password_activation
End Sub

Private Sub cmdactivar_Click()
Call activacion_temporal(Trim(Me.TxtPassword.Text))
End Sub

Private Sub cmdEliminarImpresora_Click()
Procendencia = Eliminar
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdImpresora_Click()
Dim impresora As String
impresora = Replace(Trim(Me.TxtImpresora.Text), "\", "\\")

strCadena = "INSERT INTO impresora (descripcion,id_alm,ruc)VALUES('" & impresora & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcEmpresa.BoundText & "')"
CnBd.Execute (strCadena)

Call llenar_printer(Me.HfPrinter, Me.DtcEmpresa.BoundText)
Me.Frame2.Visible = False
End Sub

Private Sub Command1_Click()
Procendencia = buscar
FrmSeguridad.Show
End Sub
Public Sub llenar_printer(ByVal Grilla As MSHFlexGrid, ByVal in_ruc As String)

strCadena = "SELECT * FROM impresora WHERE ruc='" & in_ruc & "' and id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' order by DESCRIPCION"
Call ConfiguraRstT(strCadena)
'Call Cargar_FlexGrid(Me.HfActividades, 8, rstT)

If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
                            
                           
                            ' edita la celda
                            
                            
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 600
        Next
        cabecera = "IDIMPRESORA" & vbTab & "NOMBRE IMPRESORA" & vbTab & "PRINT"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
        
        c = 2
        NumeroCampo = 2
        estado = Chr(168)
        Fila = rstT("id_impresora") & vbTab & rstT("descripcion") & vbTab & estado
          Grilla.AddItem Fila
           If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                             'If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
                          
        End If
          Fila = ""
        
          rstT.MoveNext
      Next i
   
    
     
End Sub


Private Sub Command2_Click()
Dim rmes As String * 2
Dim ranio As String * 4
Dim desVentas As String

Call put_version_update(Me.DtcEmpresa.BoundText)

If Me.chkCamara.Value = 1 Then
    KEY_CAMARA = "si"
Else
    KEY_CAMARA = "no"
End If


If Me.chkFinger.Value = 1 Then
    KEY_FINGERPRINT = "si"
Else
    KEY_FINGERPRINT = "no"
End If

If Me.chkinvitado.Value = 0 Then
 strCadena = "SELECT count(*) FROM  almacen WHERE  dni_save<>'" & KEY_USUARIO & "' AND  dni_save<>'0' AND   id_alm='" & Trim(Me.DtcVentanilla.BoundText) & "' and id_sucursal='" & Me.DtcAlmacen.BoundText & "'  AND ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 If rst(0) > 0 Then
    If MsgBox("Ventanilla Ocupada por otro Usuario del Sistema" + Chr(13) + "Para Ingresar Necesita el Password del Administrador" + Chr(13) + "Desea Continuar ?", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
       Procendencia = Selecionar
       FrmSeguridad.Show
       Exit Sub
    Else
        Me.DtcVentanilla.SetFocus
    Exit Sub
    End If
    
 End If
 End If



If Me.chkImpresora.Value = 1 Then
   KEY_IMPRESORA = "si"
    KEY_PRINTER = Trim(Me.HfPrinter.TextMatrix(Me.HfPrinter.Row, 1))
Else
    KEY_IMPRESORA = "no"
End If

   KEY_SUCURSAL = Me.DtcAlmacen.BoundText
   KEY_FECHA = Format(Me.DTPicker1.Value, "yyyy-mm-dd")

   strCadena = "SELECT * FROM gig_usuarios_online WHERE ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
        MDIFrmPrincipal.StatusBar1.Panels(4) = "USUARIOS EN LINEA:" + Space(2) + str(rst.RecordCount + 1)
    Else
        MDIFrmPrincipal.StatusBar1.Panels(4) = "USUARIOS EN LINEA:" + Space(2) + "1"
   End If
   
   
Call ActualizarControles
KEY_ALM = DtcAlmacen.BoundText
KEY_VENTANILLA = Me.DtcVentanilla.BoundText




If Me.chkinvitado.Value = 1 Then
   KEY_COMPROBANTES_PROPIOS = "no"
Else
    KEY_COMPROBANTES_PROPIOS = get_comprobante_propio(KEY_VENTANILLA)
End If



    Call put_ingreso("01")
    
    strCadena = "SELECT * FROM view_parametro_entidad WHERE cod_unico='" & Trim(Me.DtcEmpresa.BoundText) & "' LIMIT 1"
    Call ConfiguraRstK(strCadena)
    KEY_TRAMITE = rstK("tramite_documentario")
    KEY_CAJA_INDEPENDIENTE = rstK("caja_independiente")
    KEY_EMPRESA = rstK("nombre_completo")
    KEY_RUC = rstK("cod_unico")
    KEY_DIRECCION = rstK("direccion")
    KEY_SKFACTURA = rstK("factura")
    KEY_BARRAS = rstK("barras")
    KEY_CON_IGV = rstK("igv")
    KEY_COMPROBANTE = rstK("doc_cod")
    KEY_MODELO_COLOR = rstK("sub_linea_color")
    KEY_FACTURACION_ELECTRONICA = rstK("facturacion_electronica")
    KEY_IMPRESION_PROFORMA = rstK("impresion_proforma")
    KEY_RESERVA_STOCK = rstK("reserva_stock")
    
    
    KEY_PORCENTAJE_CREDITO = rstK("porcentaje_interes")
    KEY_PORCENTAJE_ZONA = rstK("porcentaje_incremento_zona")
    
    KEY_IMPUESTO_BOLSAS = rstK("impuesto_bolsas")
    KEY_VALOR_BOLSA = rstK("valor_impuesto_bolsa")
    
    KEY_DETALLE_COMBO = rstK("detalle_consumo_combo")
    
    KEY_SEGMENTACION_PRECIO = rstK("segmentacion_precio")
    KEY_RESOLUCION = rstK("resolucion_electronica")
    KEY_AUTOMATICO = rstK("automatico")
    KEY_GUIA_FRACCIONADA = rstK("guia_fraccionada")
    KEY_CERVECERIA = rstK("cerveceria")
    KEY_FOTO = rstK("foto_producto")
    KEY_PAIS = rstK("codigo_pais")
    
    KEY_SERVIDOR_CLOUD = rstK("servidor_cloud")
    KEY_SERVIDOR_KEYFACIL = rstK("servidor_keyfacil")
    KEY_TOKEN_SUCURSAL = rstK("token_sucursal")
    KEY_TOKEN_CLOUD = rstK("token")
    KEY_TOKEN_LOCAL = rstK("token_local")
    KEY_CODIGO_UNIVERSAL_IMPRESION = rstK("codigo_universal_impresion")
    KEY_MONEDA = rstK("id_moneda")
    KEY_ALERTA_CORTE = rstK("alerta_cobranza")
    KEY_SIN_EFECTO_CAJA = rstK("sin_efecto_caja")
    
    KEY_LINEA_CREDITO = rstK("linea_credito")
    KEY_FINGERPRINT = rstK("fingerprint")
    KEY_CONTABILIDAD = rstK("contabilidad")
    KEY_PROYECTO = rstK("proyectos_inversion")
    KEY_VALIDACION_EXTREMA = rstK("validacion_extrema_cliente")
    KEY_TRACKING = rstK("tracking")
    KEY_SEGURO_VENTA = rstK("servicio_seguro")
    KEY_CTA_DETRACCION = rstK("cuenta_detraccion")
    KEY_PORCENTAJE_DETRACCION = rstK("porcentaje_detraccion")
    KEY_CAMBIO_PRECIO_PASS = rstK("cambio_precio_clave")
    KEY_EMAIL = rstK("mail")
    KEY_PAQUETE_EMPRESARIAL = rstK("id_paquete_empresarial")
    KEY_ENVIO_SUNARP = rstK("envio_sunarp_xml")
    KEY_TRANSPORTE_MIGRA = rstK("transporte_integrado")
    KEY_GENERADOR_MENSUALIDAD = rstK("generador_mensualidad")
    KEY_CONTROL_MERCADERIA = rstK("control_salida_mercaderia")
    KEY_UPDATE_PROFORM = rstK("modificar_proforma")
    
    KEY_CTA_COBRAR_PRODUCTO = rstK("cuenta_cobrar_producto")
    KEY_CTA_COBRAR_SERVICIO = rstK("cuenta_cobrar_servicio")
    KEY_CTA_INGRESO_PRODUCTO = rstK("cuenta_ingreso_producto")
    KEY_CTA_INGRESO_SERVICIO = rstK("cuenta_ingreso_servicio")
    
    KEY_FECHA_CORTE = rstK("caducidad")
    
    
    KEY_CTA_PAGAR_SERVICIO = rstK("cuenta_pagar_servicio")
    KEY_CTA_IGV_VENTA = rstK("cuenta_igv_venta")
    KEY_CTA_IGV_SERVICIO_COMPRA = rstK("cuenta_igv_compra_servicio")
    
    KEY_ALARMA_STOCK = rstK("alarma_stock")
    
    KEY_GRUPO_EMPRESARIAL = rstK("grupo_empresarial")
    
    KEY_MOSTRAR_SUCURSAL = rstK("mostrar_direccion_sucursal")
    
    KEY_PRODUCTO_REPETIDO = rstK("producto_repetido")
    KEY_RUBRO = rstK("id_tipo_per")
    KEY_AGRANEL = rstK("agranel")
    KEY_DIAS_CREDITO = rstK("dias_credito")
    KEY_MORA = rstK("mora_mensualidad")
    KEY_MORA_MONTO = rstK("mora_monto")
    KEY_PROVEEDOR = rstK("id_proveedor_servicio")
    KEY_MOSTRAR_PRECIO_MAYOR = rstK("mostrar_precio_mayor")
    KEY_MOSTRAR_PRECIO_COSTO = rstK("mostrar_precio_costo")
    KEY_CTA_COMPRA_SOLES = rstK("cuenta_compra_pagar_soles")
    KEY_CTA_COMPRA_DOLARES = rstK("cuenta_compra_pagar_dolar")
    KEY_CTA_COMPRA_RH = rstK("cuenta_compra_pagar_rh")
    
    
    KEY_ASIENTO_GLOBAL_CTA_PAGAR = rstK("asiento_global_cta_pagar")
    KEY_STOCK_CONTABLE = rstK("stock_contable")
    KEY_NOTA_CREDITO_ADMIN = rstK("nota_credito_admin")
    KEY_NOTA_CREDITO_USER = rstK("nota_credito_user")
    KEY_STOCK_GLOBAL = rstK("stock_global")
    KEY_LINEA_CREDITO = rstK("linea_credito")
    KEY_GRIFO = rstK("grifo")
    KEY_REFERENCIA_COMPROBANTE = rstK("referencia_comprobante")
    KEY_BONIFICACIONES = rstK("bonificaciones")
    '***Letras pagar
    
    KEY_CTA_LETRA_PAGAR_SOLES = rstK("cuenta_letra_pagar_soles")
    KEY_CTA_LETRA_PAGAR_DOLARES = rstK("cuenta_letra_pagar_dolares")

    '*** Intrumentos de descuento FET
    
    KEY_CTA_FET_SOLES = rstK("cuenta_pagar_fet_soles")
    KEY_CTA_FET_DOLARES = rstK("cuenta_pagar_fet_dolares")
    KEY_ALERTA_CORTE = rstK("alerta_cobranza")
    KEY_TURNO = Me.DtcTurno.BoundText
    If rstK("nombre_comercial") = "-" Then
        KEY_NOMBRE_COMERCIAL = rstK("nombre_completo")
    Else
        KEY_NOMBRE_COMERCIAL = rstK("nombre_comercial")
    End If

    '****CUENTAS DE ANTICIPOS.
    KEY_CTA_ANT_SOLES = rstK("cuenta_ancipo_soles")
    KEY_CTA_ANT_DOLARES = rstK("cuenta_anticipo_dolares")
    
    
    
    
    strCadena = "SELECT descripcion FROM tipo_letra_impresion WHERE id_tipo_letra='" & rstK("id_tipo_letra") & "' LIMIT 1 "
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
       KEY_TIPO_LETRA = rstZ("descripcion")
    End If
    MDIFrmPrincipal.Caption = KEY_EMPRESA
    
    
    If KEY_ALM = "" Then
       strCadena = "SELECT id_alm,movimiento_sin_stock,direccion FROM almacen WHERE defecto='si' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
           KEY_ALM = rst(0)
           KEY_MOVIMIENTO_SIN_STOCK = rst("movimiento_sin_stock")
       Else
           KEY_ALM = "00001"
           KEY_MOVIMIENTO_SIN_STOCK = "si"
       End If
    End If
    
'Call verificacion_activacion
MDIFrmPrincipal.StatusBar1.Panels(5) = "VT N°:" & Me.DtcVentanilla.Text
MDIFrmPrincipal.StatusBar1.Panels(3) = "FECHA:" + Space(2) + Format$(KEY_FECHA, "dd-mm-yyyy")
MDIFrmPrincipal.StatusBar1.Panels(2) = "IGV:" + Space(2) + str(KEY_IGV * 100) + "%"


strCadena = "SELECT facturacion_centralizada,conversion_dolares,movimiento_sin_stock,color_barra,color,id_departamento,id_provincia,id_distrito FROM almacen WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   
   KEY_DEPARTAMENTO = rstZ("id_departamento")
   KEY_PROVINCIA = rstZ("id_provincia")
   KEY_DISTRITO = rstZ("id_distrito")
   
   KEY_COLOR_BARRA = rstZ("color_barra")
   If KEY_COLOR_BARRA = "si" Then
      MDIFrmPrincipal.StatusBar1.Panels(2).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
      MDIFrmPrincipal.StatusBar1.Panels(3).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
      MDIFrmPrincipal.StatusBar1.Panels(4).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
      MDIFrmPrincipal.StatusBar1.Panels(5).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
      MDIFrmPrincipal.StatusBar1.Panels(6).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
      'MDIFrmPrincipal.StatusBar1.Panels(7).Picture = LoadPicture(App.Path & "\Imagenes\" & rstZ("color") & ".jpg")
   End If
   
   KEY_FACTURACION_CENTRALIZADA = rstZ("facturacion_centralizada")
   KEY_CONVERSION_CAMBIO = rstZ("conversion_dolares")
   KEY_MOVIMIENTO_SIN_STOCK = rstZ("movimiento_sin_stock")
   If rstZ("color_barra") = "si" Then
        
   End If
   
   
Else
   KEY_FACTURACION_CENTRALIZADA = "si"
   KEY_CONVERSION_CAMBIO = "no"
   KEY_MOVIMIENTO_SIN_STOCK = "si"
End If


'---
If Me.chkinvitado.Value = 1 Then
    strCadena = "SELECT id_doc FROM almacen_comprobante WHERE id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' AND defecto='si' LIMIT 1"
Else

    If KEY_COMPROBANTES_PROPIOS = "si" Then
        strCadena = "SELECT id_doc FROM almacen_comprobante WHERE defecto='si' and id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' AND defecto='si' LIMIT 1"
    Else
        strCadena = "SELECT id_doc FROM almacen_comprobante WHERE id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' AND defecto='si' LIMIT 1"
    End If
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    KEY_COMPROBANTE = rst("id_doc")
Else
    MsgBox "NO HAY COMPROBANTE POR DEFECTO" + Chr(13) + "PARA ESTA SUCURSAL O VENTANILLA", vbInformation, KEY_VENDEDOR
End If

If Val(KEY_VENTANILLA) > 0 Then
    strCadena = "SELECT facturacion_detallada,comprobante_adicional FROM almacen WHERE id_alm='" & KEY_VENTANILLA & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Else
    strCadena = "SELECT facturacion_detallada,comprobante_adicional FROM almacen WHERE id_alm='" & DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' LIMIT 1"
End If
Call ConfiguraRstZ(strCadena)

If rstZ.RecordCount > 0 Then
    KEY_FACTURACION_DETALLADA = rstZ("facturacion_detallada")
    KEY_COMPROBANTE_ADICIONAL = rstZ("comprobante_adicional")
    
End If



If Me.chkinvitado.Value = 0 Then
    strCadena = "UPDATE almacen SET dni_save='" & KEY_USUARIO & "' WHERE id_alm='" & Me.DtcVentanilla.BoundText & "' and id_sucursal='" & KEY_ALM & "'  AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "UPDATE almacen SET dni_save='0' WHERE id_alm<>'" & Me.DtcVentanilla.BoundText & "' and dni_save='" & KEY_USUARIO & "' and id_sucursal='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
End If

MDIFrmPrincipal.Caption = KEY_EMPRESA & Space(3) & "[" & UCase(Trim(Me.DtcAlmacen.Text)) & "]"

If KEY_CLOUD = "si" Then
    MDIFrmPrincipal.StatusBar1.Panels(6) = "[  CLOUD  ]"
    MDIFrmPrincipal.StatusBar1.Panels(6).Picture = LoadPicture(App.Path & "\Imagenes\cloud.jpg")
Else
    MDIFrmPrincipal.StatusBar1.Panels(6) = "[AMAZON CLOUD]"
    MDIFrmPrincipal.StatusBar1.Panels(6).Picture = LoadPicture(App.Path & "\Imagenes\cloud.jpg")
End If


'If KEY_ALM <> "00001" Then
    KEY_DIRECCION_ALM = get_direccion_alm(KEY_ALM)
'lse
 '   KEY_DIRECCION_ALM = ""
'End If
 MDIFrmPrincipal.StatusBar1.Panels(1).Picture = LoadPicture(App.Path & "\archivos\menu\menu001.jpg")
 
 
 
 MDIFrmPrincipal.StatusBar1.Panels(8).Picture = LoadPicture(App.Path & "\archivos\menu\lock.jpg")
 MDIFrmPrincipal.StatusBar1.Panels(7) = "USER:" + Space(1) + KEY_VENDEDOR
If KEY_PAQUETE_EMPRESARIAL = "01" Then '    PAQUETE BASICO
    MDIFrmPrincipal.MnuMantenimientos.Visible = False
    MDIFrmPrincipal.MnuMovimientos.Visible = False
    MDIFrmPrincipal.mnucaja.Visible = False
    MDIFrmPrincipal.MnuActualizacion.Visible = False
    MDIFrmPrincipal.MnuReportes.Visible = False
    MDIFrmPrincipal.MnuInformesContables.Visible = False
    MDIFrmPrincipal.mnuInformessunat.Visible = False
    MDIFrmPrincipal.mnuInformesGerenciales.Visible = False
    MDIFrmPrincipal.MnuGestionFinanciera.Visible = False
    MDIFrmPrincipal.mnuGestionegocio.Visible = False
    MDIFrmPrincipal.MnuSeguridad.Visible = False
    frmmenu.Show
    
    Call get_telefono_sucursal(KEY_ALM)
    Unload Me
    Exit Sub
End If
If KEY_PAQUETE_EMPRESARIAL = "02" Then '    PAQUETE PROFESIONAL
End If
If KEY_PAQUETE_EMPRESARIAL = "03" Then '    PAQUETE PREMIUN
End If

 
 Call get_cambio_sbs(KEY_FECHA)
 
 
 
 
 'Call get_cambio
 Call get_telefono_sucursal(KEY_ALM)
 Unload Me
 
 
 If KEY_GRIFO = "si" Then
    frmsurtidores.Show
 Else
'    FrmDocumentos.Show
     frmmenu.Show
 End If
 
 
 If IsNull(KEY_FECHA_CORTE) = False Then
    If KEY_RUC <> "20538939618" And KEY_RUC <> "20604059136" Then
    If get_cobranza_deuda = True Then
      Call put_bloqueo
   End If
   'Call get_cobranza
   End If
  End If
  
  
  
  
  
Exit Sub
End Sub
Private Function get_cobranza_deuda() As Boolean
strCadena = "SELECT activacion_corte FROM entidad_empresa WHERE activacion_corte='si' and  cod_unico='" & KEY_RUC & "' and id_empresa='" & KEY_PROVEEDOR & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then

    strCadena = "SELECT * FROM cobranza_servicio_persona WHERE id_venta<=0 and  dni='" & KEY_RUC & "' and ruc='" & KEY_PROVEEDOR & "' LIMIT 1"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
        get_cobranza_deuda = True
    Else
        get_cobranza_deuda = False
    End If
Else
        get_cobranza_deuda = False
End If

End Function




Private Sub DtcAlmacen_Change()

Call load_ventanillas(Me.DtcAlmacen.BoundText, Me.DtcEmpresa.BoundText)


End Sub



Private Sub load_ventanillas(ByVal in_alm As String, ByVal in_ruc As String)

strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion FROM almacen where id_sucursal='" & in_alm & "' and id_tipoentidad='00012' and ruc='" & in_ruc & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcVentanilla)

End Sub
Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command2.SetFocus
End If
End Sub

Private Sub DtcEmpresa_Change()

Call llenar_almacen(Me.DtcEmpresa.BoundText)
Call get_turno(Me.DtcEmpresa.BoundText)
Call llenar_ventanilla(Me.DtcAlmacen.BoundText)
Call get_bandera

End Sub


Private Sub get_bandera()
'--------- foto--------
strCadena = "SELECT codigo_pais FROM entidad_parametros WHERE cod_unico='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then

If Len(rstA("codigo_pais")) = 4 Then
    If VerificarFichero(App.Path & "\archivos\menu") = True Then
        Me.imgPais.Picture = LoadPicture(App.Path + "\archivos\menu\" + rstA("codigo_pais") + ".jpg")
        
    Else
        Me.imgPais = Nothing
    End If
End If
End If
'--------- foto--------
End Sub
Private Sub DtcTurno_Change()
strCadena = "SELECT * FROM turno WHERE id_turno='" & Trim(Me.DtcTurno.BoundText) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblHoras.Caption = "[ " & Format(rst("hora_inicio"), "Medium Time") & "-" & Format(rst("hora_final"), "Medium Time") & " ]"
Else
    Me.lblHoras.Caption = ""
End If

End Sub
Private Sub get_turno(ByVal in_ruc As String)
Dim hi As String
Dim hf As String
Dim h As String
mostrar:
strCadena = "SELECT * FROM turno WHERE ruc='" & in_ruc & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "INSERT INTO turno (id_turno,descripcion,hora_inicio,hora_final,horas,ruc)VALUES('01','TODO EL DIA','00:01:59','23:59:59','24:00:00','" & in_ruc & "')"
    CnBd.Execute (strCadena)
    GoTo mostrar
End If

rst.MoveFirst
strCadena = "SELECT DATE_SUB(NOW(), INTERVAL 5 HOUR);"
Call ConfiguraRstT(strCadena)

h = Format(rstT(0), "hh:mm:ss")
For i = 0 To rst.RecordCount - 1
    hi = Format(rst("hora_inicio"), "hh:mm:ss")
    hf = Format(rst("hora_final"), "hh:mm:ss")
    
    If Format(TimeValue(hi), "hh:mm:ss") > Format(TimeValue(hf), "hh:mm:ss") Then
        
        hf = Format(TimeValue(hi) + TimeValue("12:00:00"), "hh:mm:ss")
        hi = Format(TimeValue(hi) + TimeValue("12:00:00"), "hh:mm:ss")
        h = Format(TimeValue(h) + TimeValue("12:00:00"), "hh:mm:ss")
        If TimeValue(h) >= TimeValue(hi) Then
            KEY_TURNO = rst("id_turno")
            Exit For
        End If
    End If
    
    
    If TimeValue(h) >= TimeValue(hi) And TimeValue(h) < TimeValue(hf) Then
        KEY_TURNO = rst("id_turno")
        Exit For
    End If
        rst.MoveNext
Next i
strCadena = "SELECT id_turno as Codigo,descripcion as Descripcion,hora_inicio FROM turno WHERE ruc='" & in_ruc & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTurno)
Me.DtcTurno.BoundText = KEY_TURNO

End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 1500
Me.DTPicker1.Value = CVDate(Date)


strCadena = "SELECT ruc as Codigo,nombre_completo as Descripcion FROM view_empresas WHERE dni='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 1 Then
    Me.TxtEmpresa.Visible = True
Else
    Me.TxtEmpresa.Visible = False
End If
Call LlenaDataCombo(Me.DtcEmpresa)
 

On Error GoTo SALIRF

strCadena = "SELECT foto,sexo,nombre_completo FROM persona WHERE dni='" & KEY_USUARIO & "' LIMIT 1"
Call ConfiguraRst(strCadena)
Me.lblUsuario.Caption = UCase(rst("nombre_completo"))
If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & KEY_USUARIO) = True Then
                sArchivo = Dir(App.Path & "\archivos\" & KEY_USUARIO & "\" & rst("foto"))
                If sArchivo <> "" Then
                     Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_USUARIO + "\" + Trim(rst("foto")))
                Else
                    GoTo sinfoto
                End If
                
              
            Else
sinfoto:
                If rst("sexo") = "M" Then
                    Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_men.jpg")
                Else
                    Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_dama.jpg")
                End If
            End If
        Else
            If rst("sexo") = "M" Then
                Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_men.jpg")
            Else
                Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_dama.jpg")
            End If
        End If
   
SALIRF:
   If rst("sexo") = "M" Then
                Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_men.jpg")
            Else
                Me.Image1.Picture = LoadPicture(App.Path + "\archivos\img_dama.jpg")
            End If
            


Me.lblDni.Caption = "DNI" + Space(1) + KEY_USUARIO














End Sub
Public Function get_nombre_cargo(ByVal in_cargo As String, ByVal in_ruc As String)
strCadena = "SELECT * FROM persona_cargos WHERE id_cargo='" & in_cargo & "' and id_empresa='" & in_ruc & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    Me.lblcargo.Caption = rstL("descripcion")
Else
    Me.lblcargo.Caption = "NO ASIGNADO"
End If
End Function
Public Function get_cargo(ByVal in_ruc As String) As String

strCadena = "SELECT id_cargo,habilitado_nota_credito FROM entidad_empresa WHERE id_empresa='" & in_ruc & "' and cod_unico='" & KEY_USUARIO & "' limit 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_cargo = rstK("id_cargo")
    KEY_HABILITADO_NOTACREDITO = rstK("habilitado_nota_credito")
    
    Me.lblcargo.Caption = get_nombre_cargo(get_cargo, in_ruc)
    If rstK("id_cargo") = "00004" Or rstK("id_cargo") = "00003" Or rstK("id_cargo") = "00009" Then
       Me.chkinvitado.Visible = True
    Else
       Me.chkinvitado.Visible = False
    End If
End If

End Function
Public Sub llenar_almacen(ByVal in_ruc As String)
KEY_CARGO = get_cargo(in_ruc)
KEY_RUC = in_ruc
Select Case KEY_CARGO
    Case "00023" ' farmacia
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  ruc='" & KEY_RUC & "'AND id_tipoentidad='00001'"
        
    Case "00001" ' admision
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM  almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
    
    Case "00008" ' admision
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
    
    Case "00006" ' asistente contable
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
    
    Case "00033" ' admision
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
        
    Case "00003"
        'Case "00001" ' Administrador
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
    Case "00004"
        'Case "00001" ' Administrador
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
    Case "00014"
        'Case "00001" ' MEDICO
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
    Case "00019"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "' AND id_alm='00001'"
    Case "00025"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "' AND id_alm='00001'"
    
    Case "00009"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
        
    Case "00015"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  ruc='" & KEY_RUC & "' AND id_alm='00001'"
        
    Case "00027"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  ruc='" & KEY_RUC & "' AND id_alm='00001'"
    
Case "00052"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
    
    Case "00058"
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
        
   Case "00057" ' admision
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM  almacen WHERE  id_tipoentidad='0' and ruc='" & KEY_RUC & "' "
    
    End Select
    
    
    
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)

End Sub
Public Sub llenar_ventanilla(ByVal in_sucursal As String)

  
   
        
        strCadena = "SELECT id_alm as Codigo,CONCAT(descripcion,`funct_estado_almacen`(dni_save,id_alm)) as Descripcion  FROM almacen WHERE id_tipoentidad='00012' and id_sucursal='" & in_sucursal & "' and ruc='" & Me.DtcEmpresa.BoundText & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcVentanilla)
        Me.DtcVentanilla.Enabled = True

End Sub


Private Sub ActualizarControles()
Dim i As Integer
Dim D As String

If KEY_CARGO = "00004" Or KEY_CARGO = "00009" Or KEY_CARGO = "00000" Then
MDIFrmPrincipal.Toolbar1.Enabled = True
MDIFrmPrincipal.MnuMantenimientos.Enabled = True
MDIFrmPrincipal.MnuActualizacion.Enabled = True
MDIFrmPrincipal.MnuMovimientos.Enabled = True
MDIFrmPrincipal.MnuReportes.Enabled = True
MDIFrmPrincipal.MnuSeguridad.Enabled = True
MDIFrmPrincipal.Toolbar1.Enabled = True
End If


If KEY_CARGO = "00057" Then
MDIFrmPrincipal.Toolbar1.Enabled = True
MDIFrmPrincipal.MnuMovimientos.Enabled = True
MDIFrmPrincipal.MnuMantenimientos.Enabled = True
MDIFrmPrincipal.MnuActualizacion.Enabled = True
MDIFrmPrincipal.mnucaja.Enabled = False
MDIFrmPrincipal.MnuReportes.Enabled = True
MDIFrmPrincipal.MnuSeguridad.Enabled = True
MDIFrmPrincipal.Toolbar1.Enabled = True
End If



If KEY_CARGO = "00008" Then
    MDIFrmPrincipal.Toolbar1.Enabled = True
    MDIFrmPrincipal.MnuMantenimientos.Enabled = True
    MDIFrmPrincipal.MnuActualizacion.Enabled = False
    MDIFrmPrincipal.MnuMovimientos.Enabled = True
    MDIFrmPrincipal.MnuActualizacion.Enabled = True
    MDIFrmPrincipal.MnuActualizarPrecio.Enabled = False
    MDIFrmPrincipal.MnuReportes.Enabled = True
    MDIFrmPrincipal.mnucaja.Enabled = True
    MDIFrmPrincipal.MnuSeguridad.Enabled = False
    MDIFrmPrincipal.MnuInformesContables = False
End If

If KEY_CARGO = "00052" Then
    MDIFrmPrincipal.Toolbar1.Enabled = True
    MDIFrmPrincipal.MnuMantenimientos.Enabled = True
    MDIFrmPrincipal.MnuActualizacion.Enabled = True
    MDIFrmPrincipal.MnuMovimientos.Enabled = True
    MDIFrmPrincipal.MnuReportes.Enabled = True
    MDIFrmPrincipal.mnucaja.Enabled = False
    MDIFrmPrincipal.MnuSeguridad.Enabled = False
    MDIFrmPrincipal.MnuInformesContables = False
    


End If

If KEY_CARGO = "00058" Then
    MDIFrmPrincipal.Toolbar1.Enabled = True
    MDIFrmPrincipal.MnuMantenimientos.Enabled = True
    MDIFrmPrincipal.MnuActualizacion.Enabled = True
    MDIFrmPrincipal.MnuMovimientos.Enabled = True
    MDIFrmPrincipal.MnuReportes.Enabled = True
    MDIFrmPrincipal.mnucaja.Enabled = False
    MDIFrmPrincipal.MnuSeguridad.Enabled = False
    MDIFrmPrincipal.MnuInformesContables = False
    MDIFrmPrincipal.mnutransferencias10.Enabled = True
    MDIFrmPrincipal.mnuordencompra.Enabled = True
    MDIFrmPrincipal.mnuordencompra10.Enabled = True
    MDIFrmPrincipal.mnuParteDiaria.Enabled = False
    MDIFrmPrincipal.MnuPagoFacturas.Visible = False
    MDIFrmPrincipal.MnuListadoDeudores01.Visible = False


End If

If KEY_CARGO = "00001" Then ' ventas
MDIFrmPrincipal.Toolbar1.Enabled = True
MDIFrmPrincipal.MnuMantenimientos.Enabled = True
MDIFrmPrincipal.MnuActualizacion.Enabled = False
MDIFrmPrincipal.MnuMovimientos.Enabled = True
MDIFrmPrincipal.MnuReportes.Enabled = True
MDIFrmPrincipal.mnucaja.Enabled = False
MDIFrmPrincipal.MnuSeguridad.Enabled = False

 
End If
If KEY_CARGO = "00006" Then ' ventas
MDIFrmPrincipal.Toolbar1.Enabled = True
MDIFrmPrincipal.MnuMantenimientos.Enabled = False
MDIFrmPrincipal.MnuActualizacion.Enabled = False
MDIFrmPrincipal.MnuMovimientos.Enabled = True
MDIFrmPrincipal.MnuReportes.Enabled = True
MDIFrmPrincipal.mnucaja.Enabled = True
MDIFrmPrincipal.MnuSeguridad.Enabled = False

 
End If

If KEY_CARGO = "00033" Then ' ventas
MDIFrmPrincipal.Toolbar1.Enabled = True
MDIFrmPrincipal.MnuMantenimientos.Enabled = False
MDIFrmPrincipal.MnuActualizacion.Enabled = False
MDIFrmPrincipal.MnuMovimientos.Enabled = True
MDIFrmPrincipal.MnuReportes.Enabled = True
MDIFrmPrincipal.mnucaja.Enabled = False
MDIFrmPrincipal.MnuSeguridad.Enabled = False

 
End If

'strCadena = "SELECT * FROM menu order by id_menu ASC"
'Call ConfiguraRst(strCadena)
'rst.MoveFirst
'For i = 0 To rst.RecordCount - 1
    
 '   strCadena = "INSERT persona_permisos(id_menu,dni,estado,ruc)VALUES('" & rst("id_menu") & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
  '  CnBd.Execute (strCadena)
   ' rst.MoveNext
'Next i
'Exit Sub

          
          If KEY_CARGO = "00004" Then 'Administrador General
            MDIFrmPrincipal.Toolbar1.Buttons(1).Visible = True
     '       MDIFrmPrincipal.Toolbar1.Buttons(2).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(5).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(9).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(11).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(13).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(15).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(17).Visible = True
          End If
          
          
          If KEY_CARGO = "00057" Then 'Administrador General
            MDIFrmPrincipal.Toolbar1.Buttons(1).Visible = True
     '       MDIFrmPrincipal.Toolbar1.Buttons(2).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(5).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(9).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(11).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(13).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(15).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(17).Visible = True
          End If
          
          
          
          
          If KEY_CARGO = "00001" Then 'ventas
            MDIFrmPrincipal.Toolbar1.Buttons(1).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(5).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(9).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(11).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(13).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(15).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(17).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(19).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(21).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(23).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(25).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(27).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(29).Enabled = False
          End If
          
          If KEY_CARGO = "00008" Or KEY_CARGO = "00052" Then 'caja
            MDIFrmPrincipal.Toolbar1.Buttons(1).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(5).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(9).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(11).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(13).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(15).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(17).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(19).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(21).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(23).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(25).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(27).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(29).Enabled = False
          End If
          If KEY_CARGO = "00033" Then 'almacen
            MDIFrmPrincipal.Toolbar1.Buttons(1).Visible = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(5).Visible = False
            MDIFrmPrincipal.Toolbar1.Buttons(7).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(9).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(11).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(13).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(15).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(17).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(19).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(21).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(23).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(25).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(27).Enabled = False
            MDIFrmPrincipal.Toolbar1.Buttons(29).Enabled = True
          End If
          
          
          If KEY_CARGO = "00006" Then 'cONATBILIDAD0003
            MDIFrmPrincipal.Toolbar1.Buttons(1).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(2).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(4).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(5).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(6).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(8).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(9).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(10).Enabled = True
          
          End If
          
          
          If KEY_CARGO = "00003" Then 'ADMINISTRADOR
            MDIFrmPrincipal.Toolbar1.Buttons(1).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(2).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(3).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(4).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(5).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(6).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(7).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(8).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(9).Enabled = True
            MDIFrmPrincipal.Toolbar1.Buttons(10).Enabled = True
          End If
          
          
          ' rst.MoveNext
    'Next i
    
'End If

End Sub

  Private Sub ActualizarCampo(ByVal id_impresora As String, ByVal Grilla As MSHFlexGrid)
    Dim cursorf As Integer
     Dim estado As String
      strCadena = "SELECT * FROM impresora WHERE  ruc='" & KEY_RUC & "' and id_alm='" & Me.DtcAlmacen.BoundText & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
            
            Grilla.TextMatrix(Grilla.Row, 2) = Chr(254)
            cursorf = Grilla.Row
            
            
            strCadena = "SELECT * FROM impresora WHERE id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                rstK.MoveFirst
                For i = 0 To rstK.RecordCount - 1
                strCadena = "UPDATE impresora set defauld='no' WHERE id_impresora='" & rstK("id_impresora") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                rstK.MoveNext
                Next i
            End If
            
           
            
            
            strCadena = "UPDATE impresora set defauld='si' WHERE id_impresora='" & Val(Grilla.TextMatrix(Grilla.Row, 0)) & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            
            For i = 1 To Grilla.Rows - 1
                If i = cursorf Then
                   
                    For j = 0 To 2
                    Grilla.col = j
                    Grilla.Row = cursorf
                    Grilla.CellBackColor = &HC0FFC0
                    Next j
                Else
                    If i <> cursor Then
                    Grilla.TextMatrix(i, 2) = Chr(168)
                    For j = 1 To 2
                        Grilla.col = j
                        Grilla.Row = i
                        Grilla.CellBackColor = &HFFFFFF
                    Next j
                    End If
                End If
                
            Next i
            Grilla.Row = cursorf
            
        End If
            
End Sub



Private Sub HfPrinter_Click()

    If Val(Me.HfPrinter.TextMatrix(Me.HfPrinter.Row, 0)) > 0 Then
        Call ActualizarCampo(Me.HfPrinter.TextMatrix(Me.HfPrinter.Row, 0), HfPrinter)
        Me.cmdEliminarImpresora.Visible = True
    Else
        Me.cmdEliminarImpresora.Visible = False
    End If

End Sub

Private Sub Timer1_Timer()
If Me.camara1.Visible = True Then
   Me.camara1.Visible = False
Else
    Me.camara1.Visible = True
End If
End Sub

Private Sub txtEmpresa_Change()

If Len(Trim(Me.TxtEmpresa.Text)) > 2 Then
strCadena = "SELECT ruc as Codigo,nombre_completo as Descripcion FROM view_empresas WHERE dni='" & KEY_USUARIO & "' and nombre_completo LIKE '%" & Trim(Me.TxtEmpresa.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEmpresa)
End If
End Sub

Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

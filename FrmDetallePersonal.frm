VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetallePersonal 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRazonSocial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   1470
      MaxLength       =   60
      TabIndex        =   169
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtIdKeyfacil 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   1485
      MaxLength       =   80
      TabIndex        =   166
      Top             =   2400
      Width           =   4935
   End
   Begin VitekeySoft.ChameleonBtn CmdFoto 
      Height          =   375
      Left            =   6720
      TabIndex        =   121
      Top             =   3280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "TOMAR INSTANTANEA"
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
      MICON           =   "FrmDetallePersonal.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSubirFoto 
      Height          =   375
      Left            =   6720
      TabIndex        =   120
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "SUBIR FOTO"
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
      MICON           =   "FrmDetallePersonal.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtRUC 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4230
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame div_verifica 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4950
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   15
         Top             =   0
         Width           =   1590
      End
   End
   Begin VB.TextBox txtMaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1470
      MaxLength       =   60
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1470
      MaxLength       =   60
      TabIndex        =   11
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtPaterno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1470
      MaxLength       =   50
      TabIndex        =   10
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox TxtDistrito 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5730
      MaxLength       =   80
      TabIndex        =   9
      Top             =   8660
      Width           =   1335
   End
   Begin VB.CommandButton cmdConSUNAT 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6030
      TabIndex        =   8
      Tag             =   "0"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   315
      Left            =   1470
      MaxLength       =   100
      TabIndex        =   7
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2250
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Txttelefono1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2370
      MaxLength       =   10
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox TxtEmail 
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
      Height          =   315
      Left            =   1470
      MaxLength       =   80
      TabIndex        =   4
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox TxtTelefono2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   -2235
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox TxtDireccion1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   1470
      MaxLength       =   150
      TabIndex        =   2
      Top             =   2060
      Width           =   4935
   End
   Begin VB.TextBox TxtDia 
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
      Height          =   300
      Left            =   1470
      MaxLength       =   11
      TabIndex        =   1
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox TxtAnio 
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
      Height          =   300
      Left            =   3765
      MaxLength       =   11
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   7230
      TabIndex        =   17
      Top             =   8160
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   10590
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":0038
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":0354
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":07B4
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":0C14
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":0F30
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":1390
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":16AC
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":1B0C
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":1F6C
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":284C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":2B68
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePersonal.frx":2E84
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet inetConecta 
      Left            =   18360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SstKardex 
      Height          =   3555
      Left            =   75
      TabIndex        =   19
      Top             =   4280
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6271
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CREDITOS"
      TabPicture(0)   =   "FrmDetallePersonal.frx":31A0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape4"
      Tab(0).Control(1)=   "lblEmpresa"
      Tab(0).Control(2)=   "Shape6"
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(4)=   "HfdPersona"
      Tab(0).Control(5)=   "chkEmpresa"
      Tab(0).Control(6)=   "ChkMaximoCredito"
      Tab(0).Control(7)=   "OptSincredito"
      Tab(0).Control(8)=   "txtRucEmpresa"
      Tab(0).Control(9)=   "txtMaximoCredito"
      Tab(0).Control(10)=   "TxtBuscar"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "EMPLEADO"
      TabPicture(1)   =   "FrmDetallePersonal.frx":31BC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label34"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label9"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label19"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label20"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label21"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label22(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label23"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label24"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label25"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label26"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "DtcEspecialidad"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "DtcSucursal"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "DtcRegimen"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "DtcPlanilla"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "DtcAfp"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "DtcCargo"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "DtpIngreso"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TxtBonificacion"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtSueldoMensual"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtAsiganacion_familiar"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TxtCuspp"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "TxtGratificacion"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "TxtRentaquinta"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "TxtSNDP"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "TxtEssalud"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "TxtPassword"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Check1"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtImagen"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "chkhabilitado"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "chk_nota_credito"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "chk_impresion"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "cmdPlanilla"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "frmPlanilla"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "TELEFONOS"
      TabPicture(2)   =   "FrmDetallePersonal.frx":31D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape7"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "HfTelefonos"
      Tab(2).Control(4)=   "DtcArea"
      Tab(2).Control(5)=   "Command1"
      Tab(2).Control(6)=   "Command3"
      Tab(2).Control(7)=   "TxtFono"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "SERVICES"
      TabPicture(3)   =   "FrmDetallePersonal.frx":31F4
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblLicencia"
      Tab(3).Control(1)=   "lblbusquedarapida"
      Tab(3).Control(2)=   "DtcEmpresaVinculada"
      Tab(3).Control(3)=   "TxtLicencia"
      Tab(3).Control(4)=   "txtBusquedaRapida"
      Tab(3).Control(5)=   "chkEmpresaVinculada"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "CTAS BANCARIAS"
      TabPicture(4)   =   "FrmDetallePersonal.frx":3210
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(1)=   "Label17"
      Tab(4).Control(2)=   "HfCuentas"
      Tab(4).Control(3)=   "DtcMoneda"
      Tab(4).Control(4)=   "DtcBanco"
      Tab(4).Control(5)=   "txtnumerocuenta"
      Tab(4).Control(6)=   "cmdagregarcuenta"
      Tab(4).Control(7)=   "cmdeliminarcuenta"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "FAMILIARES"
      TabPicture(5)   =   "FrmDetallePersonal.frx":322C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label27"
      Tab(5).Control(1)=   "Label28"
      Tab(5).Control(2)=   "Label29"
      Tab(5).Control(3)=   "Label30"
      Tab(5).Control(4)=   "Label31"
      Tab(5).Control(5)=   "Label32"
      Tab(5).Control(6)=   "DtcParentesco"
      Tab(5).Control(7)=   "HfgFamiliares"
      Tab(5).Control(8)=   "TxtFnombers"
      Tab(5).Control(9)=   "TxtFpaterno"
      Tab(5).Control(10)=   "TxtTelefono"
      Tab(5).Control(11)=   "cmdAgregar"
      Tab(5).Control(12)=   "TxtFmaterno"
      Tab(5).Control(13)=   "txtDni"
      Tab(5).ControlCount=   14
      Begin VB.Frame frmPlanilla 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2535
         Left            =   6420
         TabIndex        =   172
         Top             =   840
         Visible         =   0   'False
         Width           =   2460
         Begin VB.Frame frmnuevo 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            TabIndex        =   177
            Top             =   120
            Width           =   2240
            Begin VB.CommandButton cmdProcesarSueldo 
               Caption         =   "PROCESAR"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   840
               TabIndex        =   182
               Top             =   1200
               Width           =   1185
            End
            Begin VB.TextBox txtMontoSueldo 
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
               Left            =   840
               MaxLength       =   10
               TabIndex        =   181
               Text            =   "0.00"
               Top             =   840
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo DtcPeriodo 
               Height          =   315
               Left            =   120
               TabIndex        =   179
               Top             =   360
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
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MONTO :"
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
               Left            =   135
               TabIndex        =   180
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PERIODO"
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
               Left            =   150
               TabIndex        =   178
               Top             =   120
               Width           =   645
            End
         End
         Begin VB.CommandButton cmdSalirSueldo 
            Caption         =   "SALIR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1680
            TabIndex        =   176
            Top             =   2040
            Width           =   705
         End
         Begin VB.CommandButton cmdEliminarSueldo 
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
            Height          =   280
            Left            =   840
            TabIndex        =   175
            Top             =   2040
            Width           =   825
         End
         Begin VB.CommandButton cmdNuevoSueldo 
            Caption         =   "NUEVO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   120
            TabIndex        =   174
            Top             =   2040
            Width           =   705
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPlanilla 
            Height          =   1935
            Left            =   0
            TabIndex        =   173
            Top             =   0
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   3413
            _Version        =   393216
            ForeColor       =   8388608
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
      End
      Begin VB.CommandButton cmdPlanilla 
         Caption         =   "HISTORIAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   5560
         TabIndex        =   171
         Top             =   880
         Width           =   830
      End
      Begin VB.CheckBox chk_impresion 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "IMP DE PROFORMAS"
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
         Height          =   250
         Left            =   4440
         TabIndex        =   168
         Top             =   3090
         Width           =   1935
      End
      Begin VB.CheckBox chk_nota_credito 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "ACTIVADO NOTA CRE"
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
         Height          =   250
         Left            =   4440
         TabIndex        =   167
         Top             =   2800
         Width           =   1935
      End
      Begin VB.CheckBox chkhabilitado 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "HABILITADO"
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
         Height          =   250
         Left            =   4440
         TabIndex        =   165
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtImagen 
         Height          =   285
         Left            =   7200
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkEmpresaVinculada 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "EMPRESA VINCULADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   119
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtBusquedaRapida 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -70440
         MaxLength       =   10
         TabIndex        =   118
         Top             =   1150
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "MOSTRAR PASS"
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
         Left            =   7200
         TabIndex        =   48
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox TxtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   7200
         MaxLength       =   100
         PasswordChar    =   "*"
         TabIndex        =   47
         Text            =   "0.00"
         Top             =   1590
         Width           =   1455
      End
      Begin VB.TextBox txtDni 
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
         Left            =   -73815
         TabIndex        =   46
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TxtFmaterno 
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
         Left            =   -73815
         TabIndex        =   45
         Top             =   1740
         Width           =   1815
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AGREGAR"
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
         Left            =   -69975
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2355
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
         Left            =   -69975
         TabIndex        =   43
         Top             =   1995
         Width           =   1935
      End
      Begin VB.TextBox TxtFpaterno 
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
         Left            =   -73815
         TabIndex        =   42
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TxtFnombers 
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
         Left            =   -73815
         TabIndex        =   41
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox TxtEssalud 
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
         Left            =   8085
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox TxtSNDP 
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
         Left            =   8085
         MaxLength       =   10
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox TxtRentaquinta 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TxtGratificacion 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   1830
         Width           =   1935
      End
      Begin VB.TextBox TxtCuspp 
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   1185
         Width           =   2175
      End
      Begin VB.CommandButton cmdeliminarcuenta 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   -70440
         TabIndex        =   35
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdagregarcuenta 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   -71640
         TabIndex        =   34
         Top             =   420
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
         Left            =   -71640
         MaxLength       =   50
         TabIndex        =   33
         Top             =   780
         Width           =   2175
      End
      Begin VB.TextBox TxtLicencia 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -72720
         MaxLength       =   10
         TabIndex        =   32
         Top             =   2520
         Width           =   2055
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
         Left            =   -70080
         MaxLength       =   10
         TabIndex        =   31
         Top             =   1500
         Width           =   2775
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
         Left            =   -70080
         TabIndex        =   30
         Top             =   1920
         Width           =   1335
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
         Left            =   -68640
         TabIndex        =   29
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TxtBuscar 
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
         Left            =   -68325
         MaxLength       =   11
         TabIndex        =   28
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtAsiganacion_familiar 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtSueldoMensual 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   885
         Width           =   1095
      End
      Begin VB.TextBox TxtBonificacion 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   1515
         Width           =   1935
      End
      Begin VB.TextBox txtMaximoCredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -72480
         MaxLength       =   11
         TabIndex        =   24
         Top             =   1380
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtRucEmpresa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -72480
         MaxLength       =   11
         TabIndex        =   23
         Top             =   660
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptSincredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "SIN CREDITO "
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
         Top             =   1860
         Width           =   2055
      End
      Begin VB.OptionButton ChkMaximoCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "ASIGNAR LINEA CREDITO :"
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
         TabIndex        =   21
         Top             =   1380
         Width           =   2055
      End
      Begin VB.OptionButton chkEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "EMPRESA VINCULADA:"
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
         TabIndex        =   20
         Top             =   660
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpIngreso 
         Height          =   300
         Left            =   1320
         TabIndex        =   49
         Top             =   1500
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         Format          =   151453697
         CurrentDate     =   41134
      End
      Begin MSDataListLib.DataCombo DtcCargo 
         Height          =   315
         Left            =   4440
         TabIndex        =   50
         Top             =   540
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
      Begin MSDataListLib.DataCombo DtcAfp 
         Height          =   315
         Left            =   1320
         TabIndex        =   51
         Top             =   825
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
         Height          =   1455
         Left            =   -70845
         TabIndex        =   52
         Top             =   855
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2566
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
      Begin MSDataListLib.DataCombo DtcPlanilla 
         Height          =   315
         Left            =   1320
         TabIndex        =   53
         Top             =   480
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
      Begin MSDataListLib.DataCombo DtcArea 
         Height          =   330
         Left            =   -70080
         TabIndex        =   54
         Top             =   720
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
      Begin MSDataListLib.DataCombo DtcBanco 
         Height          =   315
         Left            =   -74070
         TabIndex        =   55
         Top             =   420
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   315
         Left            =   -74070
         TabIndex        =   56
         Top             =   765
         Width           =   2295
         _ExtentX        =   4048
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
      Begin MSDataListLib.DataCombo DtcRegimen 
         Height          =   315
         Left            =   1320
         TabIndex        =   57
         Top             =   1860
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
      Begin MSDataListLib.DataCombo DtcSucursal 
         Height          =   315
         Left            =   1320
         TabIndex        =   58
         Top             =   2620
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFamiliares 
         Height          =   1050
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   1852
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
      Begin MSDataListLib.DataCombo DtcParentesco 
         Height          =   315
         Left            =   -69975
         TabIndex        =   60
         Top             =   1635
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTelefonos 
         Height          =   2010
         Left            =   -74880
         TabIndex        =   61
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3545
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuentas 
         Height          =   1290
         Left            =   -74040
         TabIndex        =   62
         Top             =   1200
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2275
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
      Begin MSDataListLib.DataCombo DtcEspecialidad 
         Height          =   315
         Left            =   1320
         TabIndex        =   63
         Top             =   2240
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
      Begin MSDataListLib.DataCombo DtcEmpresaVinculada 
         Height          =   315
         Left            =   -72360
         TabIndex        =   116
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
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
      Begin VB.Label lblbusquedarapida 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSQUEDA RAPIDA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   -72390
         TabIndex        =   117
         Top             =   1275
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI :"
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
         Height          =   195
         Left            =   -74250
         TabIndex        =   91
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENTESCO :"
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
         Height          =   195
         Left            =   -71265
         TabIndex        =   90
         Top             =   1755
         Width           =   1125
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO :"
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
         Height          =   195
         Left            =   -71055
         TabIndex        =   89
         Top             =   1995
         Width           =   915
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.MATERNO :"
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
         Height          =   195
         Left            =   -74880
         TabIndex        =   88
         Top             =   1755
         Width           =   1035
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A.PATERNO :"
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
         Height          =   195
         Left            =   -74850
         TabIndex        =   87
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRES :"
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
         Height          =   195
         Left            =   -74700
         TabIndex        =   86
         Top             =   2355
         Width           =   855
      End
      Begin VB.Label Label26 
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
         Left            =   450
         TabIndex        =   85
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESSALUD :"
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
         Left            =   7260
         TabIndex        =   84
         Top             =   915
         Width           =   705
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SNDP:"
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
         Left            =   7515
         TabIndex        =   83
         Top             =   555
         Width           =   435
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTA 5TA :"
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
         Left            =   3810
         TabIndex        =   82
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRATI:"
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
         Left            =   3990
         TabIndex        =   81
         Top             =   1905
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderStyle     =   6  'Inside Solid
         X1              =   3600
         X2              =   3600
         Y1              =   435
         Y2              =   2500
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REGIMEN :"
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
         Left            =   510
         TabIndex        =   80
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INGRESO :"
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
         Left            =   90
         TabIndex        =   79
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSPP:"
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
         Left            =   720
         TabIndex        =   78
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   77
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   -74745
         TabIndex        =   76
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO TELEFONO"
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
         Left            =   -70125
         TabIndex        =   75
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO TELEFONICO"
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
         Left            =   -70140
         TabIndex        =   74
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label lblLicencia 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LICENCIA CONDUCIR :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   -74670
         TabIndex        =   73
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   1935
         Left            =   -70440
         Top             =   420
         Width           =   4335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR :"
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
         Left            =   -70605
         TabIndex        =   72
         Top             =   615
         Width           =   645
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   1935
         Left            =   -70920
         Top             =   420
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLANILLA :"
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
         Left            =   510
         TabIndex        =   71
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARGO :"
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
         Left            =   3870
         TabIndex        =   70
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUELDO :"
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
         Left            =   3810
         TabIndex        =   69
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFP :"
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
         Left            =   900
         TabIndex        =   68
         Top             =   780
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASIG.FAM :"
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
         Left            =   3690
         TabIndex        =   67
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BONIFICA:"
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
         Left            =   3750
         TabIndex        =   66
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label lblEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         TabIndex        =   65
         Top             =   1020
         Width           =   3615
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   -74880
         Top             =   420
         Width           =   3855
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESPECIALIDAD :"
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
         Left            =   180
         TabIndex        =   64
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3050
         Left            =   45
         Top             =   420
         Width           =   8895
      End
   End
   Begin MSDataListLib.DataCombo DtcDistrito 
      Height          =   330
      Left            =   1590
      TabIndex        =   92
      Top             =   8650
      Width           =   4095
      _ExtentX        =   7223
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
      Left            =   1590
      TabIndex        =   93
      Top             =   7900
      Width           =   4095
      _ExtentX        =   7223
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
      Left            =   1590
      TabIndex        =   94
      Top             =   8280
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSDataListLib.DataCombo Dtcmes 
      Height          =   330
      Left            =   2085
      TabIndex        =   95
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataCombo DtcSexo 
      Height          =   330
      Left            =   1470
      TabIndex        =   96
      Top             =   3120
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   9600
      TabIndex        =   123
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   15901
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ESTADO DE CUENTA"
      TabPicture(0)   =   "FrmDetallePersonal.frx":3248
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameDetalles"
      Tab(0).Control(1)=   "cmdDetallesueldo"
      Tab(0).Control(2)=   "ChameleonBtn1"
      Tab(0).Control(3)=   "txtasignacionfamiliar_sueldo"
      Tab(0).Control(4)=   "txtaniosueldo"
      Tab(0).Control(5)=   "txtsueldomensual_sueldo"
      Tab(0).Control(6)=   "txtcomisiones_sueldo"
      Tab(0).Control(7)=   "txtdescuentos_sueldo"
      Tab(0).Control(8)=   "txtretenciones_sueldo"
      Tab(0).Control(9)=   "txtsueldototal"
      Tab(0).Control(10)=   "txtsueldo_actual"
      Tab(0).Control(11)=   "txtprecio_hora"
      Tab(0).Control(12)=   "txtdias"
      Tab(0).Control(13)=   "txtHorasdiarias"
      Tab(0).Control(14)=   "txthorasmes"
      Tab(0).Control(15)=   "txtdomingos"
      Tab(0).Control(16)=   "txtsabados"
      Tab(0).Control(17)=   "HfAdelantos"
      Tab(0).Control(18)=   "DtcMesSueldo"
      Tab(0).Control(19)=   "HfComisiones"
      Tab(0).Control(20)=   "cmdReporte"
      Tab(0).Control(21)=   "cmdreportegeneral"
      Tab(0).Control(22)=   "Label36"
      Tab(0).Control(23)=   "Label35"
      Tab(0).Control(24)=   "Image2"
      Tab(0).Control(25)=   "Label37"
      Tab(0).Control(26)=   "Image3"
      Tab(0).Control(27)=   "Label38"
      Tab(0).Control(28)=   "Image4"
      Tab(0).Control(29)=   "Image5"
      Tab(0).Control(30)=   "Label40"
      Tab(0).Control(31)=   "Image6"
      Tab(0).Control(32)=   "Label41"
      Tab(0).Control(33)=   "Image8"
      Tab(0).Control(34)=   "Label43"
      Tab(0).Control(35)=   "Image9"
      Tab(0).Control(36)=   "Label44"
      Tab(0).Control(37)=   "Image10"
      Tab(0).Control(38)=   "Label45"
      Tab(0).Control(39)=   "Image11"
      Tab(0).Control(40)=   "Label39"
      Tab(0).Control(41)=   "Label42"
      Tab(0).Control(42)=   "Label46"
      Tab(0).Control(43)=   "Label47"
      Tab(0).Control(44)=   "Label48"
      Tab(0).Control(45)=   "Label49"
      Tab(0).Control(46)=   "Label50"
      Tab(0).Control(47)=   "Shape2"
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "CONTROL DE ASISTENCIA"
      TabPicture(1)   =   "FrmDetallePersonal.frx":3264
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DtcPeriodoAsistencia"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "HfAsistencia"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdconsultar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "FUNCIONES"
      TabPicture(2)   =   "FrmDetallePersonal.frx":3280
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdconsultar 
         Caption         =   "CONSULTAR"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   186
         Top             =   800
         Width           =   1215
      End
      Begin VB.Frame frameDetalles 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8415
         Left            =   -84840
         TabIndex        =   160
         Top             =   360
         Visible         =   0   'False
         Width           =   9855
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSueldo 
            Height          =   7575
            Left            =   480
            TabIndex        =   161
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   13361
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
         Begin VitekeySoft.ChameleonBtn cmdCerrar 
            Height          =   315
            Left            =   7680
            TabIndex        =   162
            Top             =   7920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "CERRAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
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
            MICON           =   "FrmDetallePersonal.frx":329C
            PICN            =   "FrmDetallePersonal.frx":32B8
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
      Begin VitekeySoft.ChameleonBtn cmdDetallesueldo 
         Height          =   315
         Left            =   -68880
         TabIndex        =   159
         Top             =   7560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "VER DETALLE"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallePersonal.frx":62CD
         PICN            =   "FrmDetallePersonal.frx":62E9
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
         Height          =   375
         Left            =   -71160
         TabIndex        =   158
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "PROCESAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetallePersonal.frx":6883
         PICN            =   "FrmDetallePersonal.frx":689F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtasignacionfamiliar_sueldo 
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
         Left            =   -70920
         TabIndex        =   138
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox txtaniosueldo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -72720
         MaxLength       =   11
         TabIndex        =   137
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtsueldomensual_sueldo 
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
         Left            =   -70920
         TabIndex        =   136
         Top             =   6000
         Width           =   1335
      End
      Begin VB.TextBox txtcomisiones_sueldo 
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
         Left            =   -70920
         TabIndex        =   135
         Text            =   "0.00"
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox txtdescuentos_sueldo 
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
         Left            =   -70920
         TabIndex        =   134
         Text            =   "0.00"
         Top             =   6720
         Width           =   1335
      End
      Begin VB.TextBox txtretenciones_sueldo 
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
         Left            =   -70920
         TabIndex        =   133
         Text            =   "0.00"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtsueldototal 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   -70920
         TabIndex        =   132
         Top             =   8160
         Width           =   1335
      End
      Begin VB.TextBox txtsueldo_actual 
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
         Left            =   -70920
         TabIndex        =   131
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox txtprecio_hora 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -68280
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   130
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox txtdias 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   129
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtHorasdiarias 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   128
         Text            =   "8.00"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txthorasmes 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -72720
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   127
         Text            =   "8.00"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtdomingos 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   126
         Text            =   "8.00"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox txtsabados 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   -69720
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   125
         Text            =   "8.00"
         Top             =   1200
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfAdelantos 
         Height          =   1335
         Left            =   -74160
         TabIndex        =   139
         Top             =   2400
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin MSDataListLib.DataCombo DtcMesSueldo 
         Height          =   315
         Left            =   -74160
         TabIndex        =   140
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfComisiones 
         Height          =   1335
         Left            =   -74160
         TabIndex        =   141
         Top             =   4080
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2355
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdReporte 
         Height          =   315
         Left            =   -68880
         TabIndex        =   163
         Top             =   7920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "VER REPORTE"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallePersonal.frx":6F71
         PICN            =   "FrmDetallePersonal.frx":6F8D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreportegeneral 
         Height          =   435
         Left            =   -68880
         TabIndex        =   164
         Top             =   8280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         BTYPE           =   5
         TX              =   "REPORTE GENERAL"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallePersonal.frx":701A
         PICN            =   "FrmDetallePersonal.frx":7036
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfAsistencia 
         Height          =   6735
         Left            =   600
         TabIndex        =   183
         Top             =   1440
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   11880
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo DtcPeriodoAsistencia 
         Height          =   330
         Left            =   1680
         TabIndex        =   185
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO  :"
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
         Height          =   210
         Left            =   690
         TabIndex        =   184
         Top             =   840
         Width           =   795
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   7935
         Left            =   240
         Top             =   480
         Width           =   9855
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "MES   :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   -74760
         TabIndex        =   157
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ADELANTOS  Y PRESTAMOS."
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
         Left            =   -74160
         TabIndex        =   156
         Top             =   2160
         Width           =   1890
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74880
         Picture         =   "FrmDetallePersonal.frx":70C3
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "COMISIONES / BONIFICACIONES"
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
         Left            =   -74160
         TabIndex        =   155
         Top             =   3840
         Width           =   2160
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74880
         Picture         =   "FrmDetallePersonal.frx":9B41
         Top             =   3720
         Width           =   480
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "ASIG. FAMILIAR :"
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
         Left            =   -73560
         TabIndex        =   154
         Top             =   5640
         Width           =   1170
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":C5BF
         Stretch         =   -1  'True
         Top             =   5640
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":CB49
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   240
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SUELDO MENSUAL :"
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
         Height          =   285
         Left            =   -73560
         TabIndex        =   153
         Top             =   6000
         Width           =   2490
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":D0D3
         Stretch         =   -1  'True
         Top             =   6360
         Width           =   240
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "COMISIONES /BONIFIC            :"
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
         Left            =   -73560
         TabIndex        =   152
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":D65D
         Stretch         =   -1  'True
         Top             =   6720
         Width           =   240
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "DESCUENTOS        :"
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
         Left            =   -73560
         TabIndex        =   151
         Top             =   6720
         Width           =   1170
      End
      Begin VB.Image Image9 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":DBE7
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   240
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "RETENCIONES :"
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
         Left            =   -73560
         TabIndex        =   150
         Top             =   7080
         Width           =   990
      End
      Begin VB.Image Image10 
         Height          =   240
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":E171
         Stretch         =   -1  'True
         Top             =   7440
         Width           =   240
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "SUELDO HASTA LA FECHA:"
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
         Left            =   -73560
         TabIndex        =   149
         Top             =   7440
         Width           =   1770
      End
      Begin VB.Image Image11 
         Height          =   360
         Left            =   -74040
         Picture         =   "FrmDetallePersonal.frx":E6FB
         Top             =   8160
         Width           =   360
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FF80&
         Caption         =   "MONTO A DEPOSITAR"
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
         Left            =   -73560
         TabIndex        =   148
         Top             =   8160
         Width           =   1470
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.HORA   :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69240
         TabIndex        =   147
         Top             =   5640
         Width           =   840
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº DIAS      :"
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
         Height          =   285
         Left            =   -74160
         TabIndex        =   146
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0C0C0&
         Caption         =   "H.DIARIAS :"
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
         Height          =   285
         Left            =   -74160
         TabIndex        =   145
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         Caption         =   "T.HORAS (MES)"
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
         Height          =   285
         Left            =   -74160
         TabIndex        =   144
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº DOMINGOS :"
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
         Height          =   285
         Left            =   -71160
         TabIndex        =   143
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº SABADOS  :"
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
         Height          =   285
         Left            =   -71160
         TabIndex        =   142
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   3255
         Left            =   -74160
         Top             =   5520
         Width           =   8535
      End
   End
   Begin SHDocVwCtl.WebBrowser wbrInfo 
      Height          =   75
      Left            =   10320
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   8640
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   132
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
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID KEYFACIL:"
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
      Height          =   210
      Index           =   1
      Left            =   345
      TabIndex        =   170
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A.MATERNO ;"
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
      Height          =   210
      Left            =   375
      TabIndex        =   115
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A.PATERNO :"
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
      Height          =   210
      Left            =   405
      TabIndex        =   114
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRES :"
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
      Height          =   210
      Left            =   525
      TabIndex        =   113
      Top             =   1320
      Width           =   795
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
      Left            =   570
      TabIndex        =   112
      Top             =   8325
      Width           =   855
   End
   Begin VB.Label lbldepartamento 
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
      Left            =   180
      TabIndex        =   111
      Top             =   7965
      Width           =   1215
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
      Left            =   690
      TabIndex        =   110
      Top             =   8685
      Width           =   705
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD UNICO:"
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
      Height          =   210
      Left            =   495
      TabIndex        =   109
      Top             =   240
      Width           =   825
   End
   Begin VB.Label LblCodPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1470
      TabIndex        =   108
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3420
      TabIndex        =   107
      Top             =   240
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   6630
      Picture         =   "FrmDetallePersonal.frx":11861
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL :"
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
      Height          =   210
      Left            =   750
      TabIndex        =   106
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label LblTelefono1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -3000
      TabIndex        =   105
      Top             =   2100
      Width           =   885
   End
   Begin VB.Label LblObservacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION:"
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
      Height          =   210
      Left            =   240
      TabIndex        =   104
      Top             =   3840
      Width           =   1065
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -1650
      TabIndex        =   103
      Top             =   2100
      Width           =   885
   End
   Begin VB.Label LblFax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   -1290
      TabIndex        =   102
      Top             =   2100
      Width           =   375
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Height          =   210
      Index           =   0
      Left            =   435
      TabIndex        =   101
      Top             =   2100
      Width           =   885
   End
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.COMPLETO :"
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
      Height          =   210
      Left            =   300
      TabIndex        =   100
      Top             =   1740
      Width           =   1005
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(dd-mm-yyyy)"
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
      Left            =   5010
      TabIndex        =   99
      Top             =   2805
      Width           =   1065
   End
   Begin VB.Label lblCumpleaños 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NACIMIENTO :"
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
      Height          =   210
      Left            =   330
      TabIndex        =   98
      Top             =   2805
      Width           =   975
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEXO :"
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
      Height          =   210
      Left            =   870
      TabIndex        =   97
      Top             =   3165
      Width           =   465
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9075
      Left            =   30
      Top             =   0
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2700
      Left            =   6630
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmDetallePersonal"
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
Public Sub llenar_datos_sueldo(ByVal ndni As String)
Dim sabados As Integer, domingos As Integer
Dim dias As Integer
Dim fecha As String
Dim horas_laborables As Single
sabados = 0
domingos = 0

For i = 1 To Val(Me.txtdias.Text)
     fecha = Format(Format(i, "00") & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & Trim(Me.txtaniosueldo.Text), "YYYY-mm-dd")
    If UCase(WeekdayName(Weekday(fecha))) = "SÁBADO" Then
       sabados = sabados + 1
    End If
    
    If UCase(WeekdayName(Weekday(fecha))) = "DOMINGO" Then
       domingos = domingos + 1
    End If
Next i

Me.txtsabados.Text = str(sabados)
Me.txtdomingos.Text = str(domingos)
dias = (Val(Me.txtdias.Text) - domingos)
horas_laborables = dias * Val(Me.txtHorasdiarias.Text)
Me.txthorasmes.Text = Format(horas_laborables, "###0.00")
Me.txtprecio_hora.Text = Format(Val(Me.txtsueldomensual_sueldo.Text) / horas_laborables, "###0.00")

strCadena = "SELECT sum(horas_trabajo) FROM persona_asistencia WHERE dni='" & Trim(Me.txtRuc.Text) & "' and fecha='" & KEY_FECHA & "' and horas_trabajo>0  "
Call ConfiguraRstZ(strCadena)
If rstZ(0) > 0 Then
    Me.txtsueldo_actual.Text = Format(Val(rstZ(0)) * Val(Me.txtprecio_hora.Text), "###0.00")
End If

End Sub

Private Sub Save()
Dim StrNombre As String
Dim StrDireccion As String
Dim fecha_nacimiento As String
Dim service As String
If Me.chkEmpresaVinculada.Value = 1 Then
    service = "si"
Else
    service = "no"
End If

If Me.chk_nota_credito.Value = 1 Then
    KEY_HABILITADO_NOTACREDITO = "si"
Else
    KEY_HABILITADO_NOTACREDITO = "no"
End If

If Me.chk_impresion.Value = 1 Then
    in_impresion = "si"
Else
    c = "no"
End If


strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico<>'" & Trim(Me.txtRuc.Text) & "' and  password='" & Trim(Me.TxtPassword.Text) & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    MsgBox "Ingrese un Password Diferente, Para evitar Errores", vbInformation
    Exit Sub
End If

Adelantado = 0
StrNombre = Comillas(Trim(Me.TxtRazonSocial.Text))
StrDireccion = Comillas(Trim(Me.TxtDireccion1.Text))
    
  If StrNombre = "" Or StrDireccion = "" Or Trim(Me.txtRuc.Text) = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
    Exit Sub
  Else
      Call verificaTipo
      strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
      Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
        'IN dni VARCHAR(11),
        'IN paterno VARCHAR(60),
        'IN materno VARCHAR(60),
        'IN nombres VARCHAR(65),
        'IN razon_social VARCHAR(150),
        'IN direccion VARCHAR(200),
        'IN mail VARCHAR(100),
        'IN id_transporte VARCHAR(2),
        'IN id_contable VARCHAR(2),
        'IN id_proveedor VARCHAR(2),
        'IN id_empleado VARCHAR(2),
        'IN id_auspiciador VARCHAR(2),
        'IN id_almacen VARCHAR(2),
        'IN ruc VARCHAR(11)
            
                strCadena = "call P_insert_persona_ii('" & Trim(Me.txtRuc.Text) & "' " & _
                ",'" & Me.txtPaterno.Text & "', " & _
                "'" & Me.txtMaterno.Text & "' " & _
                ",'" & Trim(Me.txtNombre.Text) & "' " & _
                ",'" & Trim(Me.TxtRazonSocial.Text) & "' " & _
                ",'" & Trim(Me.TxtDireccion1.Text) & "' " & _
                ",'" & Trim(Me.txtTelefono.Text) & "' " & _
                ",'" & Me.txtEmail.Text & "'" & _
                ",'" & StrTransporte & "' " & _
                ",'" & StrContable & "'" & _
                ",'" & strProveedor & "' " & _
                ",'" & StrPersonal & "' " & _
                ",'" & StrAuspiciador & "' " & _
                ",'" & StrAlmacen & "' " & _
                ",'no' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                 
                
                
                If StrAlmacen = "si" Then
                    Call persona_almacen(Trim(Me.txtRuc.Text))
                End If
                If img <> vbNullString And Trim(str(Me.txtRuc.Text)) <> vbNullString Then
                    strCadena = "UPDATE persona SET foto='" & Trim(Me.txtimagen.Text) & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                    CnBd.Execute (strCadena)
                
                     
                    img = ""
                End If
                
                If service = "si" Then
                    strCadena = "UPDATE entidad_empresa SET habilitado_nota_credito='" & KEY_HABILITADO_NOTACREDITO & "',id_personal='si' , id_especialidad='" & Me.DtcEspecialidad.BoundText & "',id_empresa_rel='" & Trim(Me.DtcEmpresaVinculada.BoundText) & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                Else
                    strCadena = "UPDATE entidad_empresa SET habilitado_nota_credito='" & KEY_HABILITADO_NOTACREDITO & "',id_personal='si' , id_especialidad='" & Me.DtcEspecialidad.BoundText & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                End If
                
                CnBd.Execute (strCadena)
                
                 
                strCadena = "UPDATE persona SET id_keyfacil='" & Trim(Me.txtIdKeyfacil.Text) & "', sexo='" & Me.DtcSexo.BoundText & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                CnBd.Execute (strCadena)
                
                 
            Else
                 cPersona = Trim(Me.LblCodPersona.Caption)
                 strCadena = "CALL PER_TrabajadorVitekey_IAE('" & KEY_RUC & "','" & cPersona & "')"
                 CnBd.Execute (strCadena)
                 
                 strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & cPersona & "'"
                 Call ConfiguraRstT(strCadena)
                 If rstT.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa)VALUES ('" & cPersona & "','" & KEY_RUC & "')"
                 Else
                    
                    If service = "si" Then
                        strCadena = "UPDATE entidad_empresa SET id_personal='si',id_cargo='" & Me.DtcCargo.BoundText & "',id_empresa_rel='" & Trim(Me.DtcEmpresaVinculada.BoundText) & "' WHERE cod_unico='" & cPersona & "' AND id_empresa='" & KEY_RUC & "'"
                    Else
                        strCadena = "UPDATE entidad_empresa SET id_personal='si',id_cargo='" & Me.DtcCargo.BoundText & "',id_especialidad='" & Me.DtcEspecialidad.BoundText & "' WHERE cod_unico='" & cPersona & "' AND id_empresa='" & KEY_RUC & "'"
                    End If
                 End If
                 
                CnBd.Execute (strCadena)
                
                 
              
                 
            End If
           ' If img <> vbNullString And Trim(Str(Me.txtRUC.text)) <> vbNullString Then
                'Ret = Guardar_Imagen(CnBd, "SELECT per_foto From Persona Where cPersona=" & Trim(cPersona), "per_foto", img)
             
            'End If
             
                If Me.DtcCargo.BoundText = "00005" Then
                   strCadena = "call SEG_Usuario_Vitekey('" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "','" & KEY_USUARIO & "','" & Trim(Me.TxtPassword.Text) & "')"
                   CnBd.Execute (strCadena)
                End If
                
                  
                    strCadena = "UPDATE persona SET id_keyfacil='" & Trim(Me.txtIdKeyfacil.Text) & "',a_paterno='" & Trim(Me.txtPaterno.Text) & "',a_materno='" & Trim(Me.txtMaterno.Text) & "',nombre_completo ='" & UCase(Trim(Me.TxtRazonSocial.Text)) & "',id_distrito='" & Me.DtcDistrito.BoundText & "',direccion='" & Trim(Me.TxtDireccion1.Text) & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_departamento='" & Me.DtcDepartamento.BoundText & "',licencia='" & Me.TxtLicencia.Text & "',id_dia='" & Trim(Me.txtdia.Text) & "',id_mes='" & Me.DtcMes.BoundText & "',id_anio='" & Trim(Me.txtAnio.Text) & "',foto='" & Trim(Me.txtimagen.Text) & "' WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
                    CnBd.Execute (strCadena)
                    
                     
             
                    If Me.ChkMaximoCredito.Value = 1 Then
                        strcredito = "si"
                    Else
                        strcredito = "no"
                    End If
                     
                     If Me.chkhabilitado.Value = 1 Then
                        in_habilitado = "si"
                     Else
                        in_habilitado = "no"
                     End If
                     
                     
                     strCadena = "UPDATE entidad_empresa SET id_especialidad='" & Me.DtcEspecialidad.BoundText & "',impresion_proforma='" & in_impresion & "',habilitado_nota_credito='" & KEY_HABILITADO_NOTACREDITO & "',habilitado='" & in_habilitado & "', id_credito='" & strcredito & "',monto_credito='" & Val(Me.txtMaximoCredito.Text) & "',id_personal='si',id_cargo='" & Me.DtcCargo.BoundText & "'," & _
                     "id_personal='" & StrPersonal & "',password='" & Trim(Me.TxtPassword.Text) & "',passwordaccesso='" & Trim(Me.TxtPassword.Text) & "',id_planilla='" & Me.DtcPlanilla.Text & "',id_sucursal='" & Me.DtcSucursal.BoundText & "',sueldo='" & Val(Me.txtSueldoMensual.Text) & "',id_proveedor='" & strProveedor & "',id_afp='" & Me.DtcAfp.BoundText & "',rta_quinta='" & Val(Me.TxtRentaquinta.Text) & "',asig_familiar='" & Val(Me.txtAsiganacion_familiar.Text) & "',bonificacion_extraordinaria='" & Val(Me.TxtBonificacion.Text) & "',cuspp='" & Val(Me.TxtCuspp.Text) & "',essalud='" & Val(Me.TxtEssalud.Text) & "',sndp='" & Val(Me.TxtSNDP.Text) & "',fecha_ingreso='" & Format(Me.DtpIngreso.Value, "YYYY-mm-dd") & "',id_cargo='" & Me.DtcCargo.BoundText & "',id_condicion='" & Me.DtcRegimen.BoundText & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "'"
                     CnBd.Execute (strCadena)
                     
                      
                    
                
                
                
                
             
            
                
                
                
                
               
    End If
       
     If FrmPersona.Procedencia = modificar Then
        Call FrmPersona.actualizar
        FrmPersona.Procedencia = Neutro
        Unload Me
        Exit Sub
     End If
     
      If frmpersonal.Procedencia = modificar Then
        Call frmpersonal.actualizar
        frmpersonal.Procedencia = Neutro
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
     If frmpersonal.Procedencia = nuevo Then
        Call frmpersonal.actualizar
        frmpersonal.Procedencia = Neutro
        Call Resalta(frmpersonal.txtRuc)
        Unload Me
        Exit Sub
     End If
     
     If FrmEspecialistas.Procedencia = modificar Then
        Call FrmEspecialistas.actualizar
        FrmEspecialistas.Procedencia = Neutro
        Unload Me
        Exit Sub
     End If
     
     
     
     If FrmEspecialistas.Procedencia = nuevo Then
        Call FrmEspecialistas.actualizar
        FrmEspecialistas.Procedencia = Neutro
        Call Resalta(FrmEspecialistas.txtRuc)
        Unload Me
        Exit Sub
     End If
       
       
       If FrmVentas.Procedencia = nuevo Or FrmVentas.Procedencia = Selecionar Then
            If Len(Trim(Me.txtRuc.Text)) = 11 Then
                FrmVentas.TxtCodCliente.Text = Trim(Me.txtRuc.Text)
                FrmVentas.Procedencia = Neutro
                
            End If
            FrmVentas.TxtCliente.Text = Trim(Me.TxtRazonSocial.Text)
            FrmVentas.TxtDireccion.Text = Trim(Me.TxtDireccion1.Text)
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
            FrmCompras.TxtDireccion.Text = Trim(Me.TxtDireccion1.Text)
            
            
        End If
        
        If frmNuevoComprobante.Procedencia = nuevo Then
            'frmNuevoComprobante.txtcodUuario.Text = Me.TxtRuc.Text
            'frmNuevoComprobante.txtCliente.Text = Me.TxtRazonSocial.Text
            'Call frmNuevoComprobante.DtcCCostos.SetFocus
        End If
        Unload Me
  

End Sub
Private Sub quitar_almacen(ByVal dni As String)
'strCadena = "DELETE FROM almacen WHERE id_responsable='" & dni & "' AND ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)
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






Private Sub ChameleonBtn2_Click()

End Sub

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.TxtPassword.PasswordChar = ""
Else
    Me.TxtPassword.PasswordChar = "*"
End If
End Sub



Private Sub Check2_Click()

End Sub

Private Sub chkEmpresa_Click()
If Me.chkEmpresa.Value = True Then
    Me.txtRucEmpresa.Visible = True
    
    Me.txtMaximoCredito.Visible = False
    
Else
  
   Me.txtRucEmpresa.Visible = False
   
   
End If
End Sub

Private Sub chkEmpresaVinculada_Click()
If Me.chkEmpresaVinculada.Value = 1 Then
      strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND id_almacen='si' ORDER BY P.nombre_completo LIMIT 0,20 "
      Call ConfiguraRst(strCadena)
      Call LlenaDataCombo(Me.DtcEmpresaVinculada)
      Me.DtcEmpresaVinculada.Enabled = True
      Me.lblbusquedarapida.Visible = True
      Me.TxtBusquedaRapida.Visible = True
      Call Resalta(Me.TxtBusquedaRapida)
      Exit Sub
Else
    Me.DtcEmpresaVinculada.Enabled = False
    Me.lblbusquedarapida.Visible = False
      Me.TxtBusquedaRapida.Visible = False
      
      Exit Sub
End If
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






Private Sub cmdagregar_Click()
Dim razon As String
If Me.txtDni.Text <> "" Then
    strCadena = "SELECT * FROM persona_accidentes WHERE dni_familia='" & Trim(Me.txtDni.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "UPDATE persona_accidentes SET telefono='" & Me.txtTelefono.Text & "',id_parentesco='" & Me.dtcparentesco.BoundText & "' WHERE dni_familia='" & Me.txtDni.Text & "' AND dni='" & Me.txtRuc.Text & "'"
        CnBd.Execute (strCadena)
         
         
    Else
        strCadena = "INSERT INTO persona_accidentes(dni,dni_familia,id_parentesco,telefono)VALUES('" & Me.txtRuc.Text & "','" & Me.txtDni.Text & "','" & Me.dtcparentesco.BoundText & "','" & Me.txtTelefono.Text & "')"
        CnBd.Execute (strCadena)
         
         
        strCadena = "select * from persona where dni='" & Trim(Me.txtDni.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
        razon = Me.TxtFnombers.Text + Space(1) + Me.TxtFpaterno.Text + Space(1) + Me.TxtFmaterno.Text
        strCadena = "P_insert_persona('" & Trim(Me.txtDni.Text) & "','" & Me.TxtFpaterno.Text & "','" & Me.TxtFmaterno.Text & "','" & Me.TxtFnombers.Text & "','" & Trim(razon) & "','CHICLAYO','" & Me.txtTelefono.Text & "','--','no','no','no','no','no','0','')"
        CnBd.Execute (strCadena)
         
         
        End If
    End If
    strCadena = "SELECT F.id,P.nombre_completo,PR.descripcion as parentesco,F.telefono,F.dni_familia FROM persona_accidentes F,persona P,parentesco PR WHERE F.id_parentesco=PR.id_parentesco AND  F.dni='" & Me.txtRuc.Text & "' AND F.dni_familia=P.dni"
    Call llenarFamiliares(Me.HfgFamiliares)
End If
End Sub

Private Sub cmdagregarcuenta_Click()
If Me.txtnumerocuenta.Text <> "" And Trim(Me.DtcBanco.BoundText) <> "" And Trim(Me.DtcMoneda.BoundText) Then
    strCadena = "INSERT INTO persona_cuentabancaria(dni,id_banco,id_moneda,cuenta)VALUES('" & Trim(Me.txtRuc.Text) & "','" & Me.DtcBanco.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Trim(Me.txtnumerocuenta.Text) & "')"
    CnBd.Execute (strCadena)
     
    Call llenarCuentas(Me.HfCuentas, Trim(Me.txtRuc.Text))
End If
End Sub

Private Sub cmdCerrar_Click()
Me.frameDetalles.Visible = False
End Sub

Private Sub cmdConsultar_Click()
Call llenar_registro(Me.HfAsistencia)



End Sub

Private Sub cmdConSUNAT_Click()
 Call precionar
End Sub

Public Sub precionar()
If cmdConSUNAT.Tag = 0 Then
        If txtRuc.Tag <> txtRuc.Text Then
            txtRuc.Tag = Trim(txtRuc.Text)
            Call CargaSunat
        End If
        
       
        cmdConSUNAT.Tag = 1
        cmdConSUNAT.Caption = "OK"
    Else
     '   Me.Width = 9195
        cmdConSUNAT.Tag = 0
        cmdConSUNAT.Caption = "OK"
    End If
    
End Sub
    
Private Sub CargaSunat()
    'cSunat.ruc = txtRuc.Text
     
    'cSunat.CargaWebExplorador
    
End Sub



Private Sub cmdDetallesueldo_Click()
Call llenar_detalle(Me.HfSueldo)
frameDetalles.Visible = True

End Sub
Public Sub llenar_detalle(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim fecha_ini As String
Dim fecha_fin As String
Dim nhoras As Single
fecha_ini = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & "01"
fecha_fin = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & Trim(Me.txtdias.Text)
Dim acumulado_horas As Single
strCadena = "SELECT * FROM persona_asistencia WHERE ruc='" & KEY_RUC & "' and dni='" & Trim(Me.txtRuc.Text) & "' and fecha>='" & Format(fecha_ini, "YYYY-mm-dd") & "'  and fecha<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY id"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1500
           
        Next
         
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA MARCACION" & vbTab & "ACCION" & vbTab & "H.TRABAJADAS" & vbTab & "VALOR EN SOLES"
        Grilla.AddItem cabecera
         For k = 1 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        acumulado_horas = 0
        acumulado_monto = 0
        rst.MoveFirst
        Grilla.ColAlignment(5) = 7
        For i = 0 To rst.RecordCount - 1
              If rst("id_acceso") = "01" Then
                nacceso = "Hora Ingreso"
              Else
                nacceso = "Hora Salida"
              End If
              
              Fila = rst("id") & vbTab & rst("fecha") & vbTab & Format(rst("hora"), "HH:mm:ss am/pm") & vbTab & nacceso & vbTab & Format(rst("horas_trabajo"), "###0.00") & vbTab & "S/. " & Format(rst("horas_trabajo") * Val(Me.txtprecio_hora.Text), "###0.00")
              Grilla.AddItem Fila
              acumulado_horas = acumulado_horas + rst("horas_trabajo")
              acumulado_monto = acumulado_monto + rst("horas_trabajo") * Val(Me.txtprecio_hora.Text)
          
         
          rst.MoveNext
             
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL HORAS:" & vbTab & Format(acumulado_horas, "###0.00") & vbTab & "S/. " & Format(acumulado_monto, "###0.00")
        Grilla.AddItem Fila
         For k = 3 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HDFDFE0
        Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Public Sub llenar_registro(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "call ADM_auditoria_empresa('8','','" & Trim(Me.txtRuc.Text) & "','" & Me.DtcPeriodoAsistencia.BoundText & "','','','','','" & KEY_RUC & "')"
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1500
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 3500
          
           
        Next
         
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA MARCACION" & vbTab & "ACCION"
        Grilla.AddItem cabecera
         For k = 1 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
      
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
        
              
              Fila = rst("id_detalle") & vbTab & rst("fecha") & vbTab & Format(rst("hora"), "HH:mm:ss am/pm") & vbTab & rst("ingreso_salida")
              Grilla.AddItem Fila
        
          
         
          rst.MoveNext
             
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL HORAS:" & vbTab & Format(acumulado_horas, "###0.00") & vbTab & "S/. " & Format(acumulado_monto, "###0.00")
        Grilla.AddItem Fila
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
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

Private Sub cmdEliminarSueldo_Click()

strCadena = "DELETE FROM persona_planilla WHERE id='" & Val(Me.HfPlanilla.TextMatrix(Me.HfPlanilla.Row, 0)) & "'"
CnBd.Execute (strCadena)
Call Me.load_planilla(Me.HfPlanilla)
End Sub

Private Sub cmdNuevoSueldo_Click()
Me.frmnuevo.Visible = True

strCadena = "SELECT id as Codigo, CONCAT(nombre,' - ',Ejercicio) as Descripcion FROM con_periodo ORDER BY Ejercicio ASC,Mes ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)

End Sub

Private Sub cmdPlanilla_Click()
Call load_planilla(Me.HfPlanilla)
Me.frmnuevo.Visible = False
Me.frmPlanilla.Visible = True

End Sub

Private Sub cmdProcesarSueldo_Click()

strCadena = "INSERT INTO persona_planilla(`id_periodo`,`dni`,`sueldo`,`ruc`)VALUES('" & Me.DtcPeriodo.BoundText & "','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtMontoSueldo.Text) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Me.frmnuevo.Visible = False


Call Me.load_planilla(Me.HfPlanilla)

End Sub
Public Sub load_planilla(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

strCadena = "CALL ADM_reportes_generales('53','','','" & Trim(Me.txtRuc.Text) & "','0','','','','','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1300
            Grilla.ColWidth(2) = 900
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
            Fila = rst("id") & vbTab & rst("nombre") & vbTab & rst("sueldo")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
     
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub cmdReporte_Click()
Dim fecha_ini As String
Dim fecha_fin As String
Dim nhoras As Single
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
fecha_ini = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & "01"
fecha_fin = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & Trim(Me.txtdias.Text)

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(2, 1) = "p_precio_hora"


arr(0, 2) = fecha_ini
arr(1, 2) = fecha_fin
arr(2, 2) = str(Me.txtprecio_hora.Text)

param = arr()


strCadena = "SELECT  dni,fecha,hora,horas_trabajo,id_acceso,nombre_completo FROM view_reporte_asistencia WHERE ruc='" & KEY_RUC & "' and dni='" & Trim(Me.txtRuc.Text) & "' and fecha>='" & Format(fecha_ini, "YYYY-mm-dd") & "'  and fecha<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY id"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_asistencia", param, App.Path + "\Reportes\")
 
End Sub

Private Sub cmdreportegeneral_Click()
Dim fecha_ini As String
Dim fecha_fin As String
Dim nhoras As Single
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
fecha_ini = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & "01"
fecha_fin = Trim(Me.txtaniosueldo.Text) & "-" & Format(Me.DtcMesSueldo.BoundText, "00") & "-" & Trim(Me.txtdias.Text)

arr(0, 1) = "p_fecha_inicio"
arr(1, 1) = "p_fecha_final"
arr(2, 1) = "p_precio_hora"


arr(0, 2) = fecha_ini
arr(1, 2) = fecha_fin
arr(2, 2) = str(Me.txtprecio_hora.Text)

param = arr()


strCadena = "SELECT  dni,fecha,hora,horas_trabajo,id_acceso,nombre_completo FROM view_reporte_asistencia WHERE ruc='" & KEY_RUC & "' and fecha>='" & Format(fecha_ini, "YYYY-mm-dd") & "'  and fecha<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY id"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_asistencia", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdSalirSueldo_Click()
Me.frmPlanilla.Visible = False

End Sub

Private Sub cmdSubirFoto_Click()
On Error GoTo finish
Dim nombre_foto As String
Dim ruta As String
Dim numero As Integer
Me.CommonDialog1.Filter = "*.Jpg"
Me.CommonDialog1.ShowOpen
Me.Image1.Picture = LoadPicture(Me.CommonDialog1.FileName)
img = Me.CommonDialog1.FileName
strCadena = "SELECT foto FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
Call ConfiguraRst(strCadena)

nombre_foto = Trim(Me.txtRuc.Text) & ".jpg"
ruta = App.Path & "\archivos\" & Trim(Me.txtRuc.Text)
    
    If VerificarFichero(ruta) = False Then
       Call MkDir(App.Path & "\archivos\" & Trim(Me.txtRuc.Text))
    End If
    
Call FileCopy(CommonDialog1.FileName, ruta & "\" & nombre_foto)
Me.txtimagen.Text = Trim(nombre_foto)
  

   
 
 
Exit Sub
finish: MsgBox "La Imagen que Intenta Subir tiene que ser .JPG", vbInformation, "Imagen no Compatible"
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
        strCadena = "INSERT INTO persona_telefono (dni,telefono,id_cargo)VALUES ('" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtFono.Text) & "','" & Me.DtcArea.BoundText & "')"
        CnBd.Execute (strCadena)
         
        
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
Public Sub llenarCuentas(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
strCadena = "SELECT abreviatura,M.descripcion,cuenta FROM banco B,moneda M,persona_cuentabancaria C WHERE B.id_banco=C.id_banco AND C.id_moneda=M.id_moneda AND C.dni='" & cPersona & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
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


Private Sub Command2_Click()

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

Private Sub Command4_Click()

End Sub

Private Sub Command6_Click()
Me.frmPlanilla.Visible = False

End Sub

Private Sub cSunat_DatosObtenidos()
    Dim Confirm As Integer
    
    If cSunat.EmRazSocial <> vbNullString Then
        If TxtRazonSocial.Text = vbNullString Then
            TxtRazonSocial.Text = cSunat.EmRazSocial
            Me.TxtDireccion1.Text = cSunat.EmDireccion
            If frmpersonal.Procedencia = nuevo Then
                If Len(Trim(frmpersonal.txtRuc.Text)) = 8 Then
                    Me.txtRuc.Text = Mid(Trim(Me.txtRuc.Text), 3, 8)
                    Me.LblCodPersona.Caption = Trim(Me.txtRuc.Text)
                End If
                
            Else
                
            End If
            'txtNomComercial.Text = cSunat.EmNomComercial
        Else
            If TxtRazonSocial.Text <> cSunat.EmRazSocial Then
                Confirm = MsgBox("La razón social que tiene almacenado el sistema no coincide con la informacion de SUNAT, ¿desea actualizar?", vbYesNo, "Confirmar actualización")
                If Confirm = vbYes Then
                    TxtRazonSocial.Text = cSunat.EmRazSocial
                     Me.TxtDireccion1.Text = cSunat.EmDireccion
                     
                End If
            End If
            
           ' If txtNomComercial.Text = vbNullString Then
            '    txtNomComercial.Text = cSunat.EmNomComercial
           ' End If
            
        End If
        
        'cmdGuardar.SetFocus
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
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        strCadena = "SELECT id_provincia FROM distrito WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            Me.lblprovincia.Visible = True
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

Private Sub Dtcmes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtAnio)
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
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcDepartamento)
        Set rst = Nothing
        Me.DtcDepartamento.Enabled = True
    End If
    Set rstTemporal = Nothing
End If
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
Me.Top = 10
Me.DtpIngreso.Value = KEY_FECHA


  strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM mes ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMes)
  
  strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM mes ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMesSueldo)
  
  Me.txtaniosueldo.Text = Year(KEY_FECHA)
  Me.DtcMesSueldo.BoundText = Month(KEY_FECHA)
  
  Me.txtdias.Text = Day(DateSerial(Me.txtaniosueldo.Text, Val(Me.DtcMesSueldo.BoundText) + 1, 0))
  
  strCadena = "SELECT id_parentesco as Codigo,descripcion as Descripcion FROM parentesco ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcparentesco)
  
  strCadena = "SELECT id_sexo as Codigo,descripcion as Descripcion FROM sexo ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcSexo)
              
  strCadena = "SELECT id_cargo as Codigo, descripcion as Descripcion FROM persona_cargos WHERE ruc='no' and id_empresa='0' ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcArea)
  
  
        
    



Select Case frmpersonal.Procedencia
        Case nuevo
        Me.OptSincredito.Value = 1
        Me.lbldepartamento.Visible = False
        Me.DtcDepartamento.Visible = False
        Me.lblprovincia.Visible = False
        Me.DtcProvincia.Visible = False
        Call llenaplanilla
        
    Case modificar
      Call llenaplanilla
      Call LLENA(Trim(frmpersonal.HfdPersona.TextMatrix(frmpersonal.HfdPersona.Row, 0)))
      
      'If Len(Me.txtRUC.Text) > 8 Then
      'Call precionar
      'End If
      
  End Select
End Sub
Private Sub llenaplanilla()
        strCadena = "SELECT id_planilla as Codigo, descripcion as Descripcion FROM planilla ORDER BY descripcion"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcPlanilla)
          If Me.DtcCargo.BoundText <> "" Then
            strCarg = Me.DtcCargo.BoundText
          End If
          strCadena = "SELECT id_cargo as Codigo, descripcion as Descripcion FROM persona_cargos WHERE ruc='si' AND id_empresa='" & KEY_RUC & "' ORDER BY descripcion"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcCargo)
          Me.DtcCargo.BoundText = strCarg
          
          strCadena = "SELECT id_afp as Codigo, descripcion as Descripcion FROM afp ORDER BY descripcion"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcAfp)
          strCadena = "SELECT id_banco as Codigo,abreviatura as Descripcion FROM banco ORDER BY abreviatura"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcBanco)
          strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcMoneda)
          strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcSucursal)
          strCadena = "SELECT id_condicion as Codigo,descripcion as Descripcion FROM condicion_laboral ORDER BY descripcion"
          Call ConfiguraRstT(strCadena)
          Call LlenaDataComboT(Me.DtcRegimen)
          
          
        
       
          
          strCadena = "SELECT id_especialidad as Codigo,descripcion as Descripcion FROM especialidad WHERE ruc='0' ORDER BY descripcion"
          Call ConfiguraRst(strCadena)
          Call LlenaDataCombo(Me.DtcEspecialidad)

End Sub
Private Sub LLENA(ByVal cPersona As String)
'On Error GoTo salir
Dim cDepartamento As String, cProvincia As String, cDistrito As String, cUrbanizacion As Double, cZona As Double
'strCadena = "SELECT * FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND P.dni = '" & FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0) & "'"
strCadena = "SELECT * FROM view_entidad WHERE ruc='" & KEY_RUC & "' AND dni = '" & Trim(cPersona) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Exit Sub
End If
  StrCodTabla = rst("dni")
  If Len(StrCodTabla) = 8 Then
     Me.SstKardex.TabVisible(5) = True
  Else
    Me.SstKardex.TabVisible(5) = False
  End If
  Me.LblCodPersona.Caption = StrCodTabla
  If IsNull(rst("dni")) = True Then
    GoTo sin_doc
  End If
  Me.txtRuc.Text = rst("dni")
  Me.txtPaterno.Text = UCase(rst("a_paterno"))
  Me.txtMaterno.Text = UCase(rst("a_materno"))
  Me.txtNombre.Text = UCase(rst("nombres"))
    
  
  If rst("habilitado") = "si" Then
     Me.chkhabilitado.Value = 1
    Else
     Me.chkhabilitado.Value = 0
  End If
    
    
  If rst("impresion_proforma") = "si" Then
     Me.chk_impresion.Value = 1
  Else
     Me.chk_impresion.Value = 0
  End If
  
  If rst("habilitado_nota_credito") = "si" Then
    Me.chk_nota_credito.Value = 1
  Else
    Me.chk_nota_credito.Value = 0
  End If
  
  
  Me.txtIdKeyfacil.Text = rst("id_keyfacil")
  Me.TxtLicencia.Text = rst("licencia")
  If Len(rst("dni")) = 8 Then
         Me.LblEntidad.Caption = "Nombre:"
         Me.LblTipoDocumento.Caption = "DNI"
         'Call llenarFamiliares(Me.HfgFamiliares)
    Else
        Me.LblEntidad.Caption = "Razon Social:"
        Me.LblTipoDocumento.Caption = "RUC"
    End If
sin_doc:
    Me.TxtRazonSocial.Text = rst("nombre_completo")
    Me.DtcSexo.BoundText = rst("sexo")




          Call llenarCuentas(Me.HfCuentas, Trim(Me.LblCodPersona.Caption))
          If IsNull(rst("password")) = False Then
            Me.TxtPassword.Text = rst("password")
           End If
          
  If rst("id_personal") = "si" Then
        
        StrPersonal = "si"
        Me.SstKardex.TabVisible(1) = True
        Me.DtcCargo.BoundText = rst("id_cargo")
        Me.DtcAfp.BoundText = rst("id_afp")
        Me.DtcSucursal.BoundText = rst("id_sucursal")
        Me.DtcPlanilla.Text = Trim(UCase(rst("id_planilla")))
        Me.txtSueldoMensual.Text = Format(rst("sueldo"), "###0.00")
       
        Me.txtsueldomensual_sueldo.Text = Format(rst("sueldo"), "###0.00")
        Me.txtasignacionfamiliar_sueldo.Text = Format(rst("asig_familiar"), "#,##0.00")
        Me.txtAsiganacion_familiar.Text = Format(rst("asig_familiar"), "#,##0.00")
        Me.TxtRentaquinta.Text = Format(rst("rta_quinta"), "#,##0.00")
        Me.DtcRegimen.BoundText = rst("id_condicion")
        Me.TxtCuspp.Text = Format(rst("cuspp"), "#,##0.00")
        Me.TxtBonificacion.Text = Format(rst("bonificacion_extraordinaria"), "#,##0.00")
        Me.TxtEssalud.Text = Format(rst("essalud"), "#,##0.00")
        Me.TxtSNDP.Text = Format(rst("sndp"), "#,##0.00")
        Me.DtcEspecialidad.BoundText = rst("id_especialidad")
        Call llenar_datos_sueldo(Trim(Me.txtRuc.Text))
        
        If IsNull(rst("fecha_ingreso")) = False Then
            Me.DtpIngreso.Value = rst("fecha_ingreso")
        Else
            Me.DtpIngreso.Value = KEY_FECHA
        End If
        
        
End If
    Me.lblCumpleaños.Visible = True
    Me.txtdia.Visible = True
    Me.DtcMes.Visible = True
    Me.txtdia.Text = formato_item(rst("id_dia"), 2)
    Me.txtAnio.Text = rst("id_anio")
    Me.DtcMes.BoundText = rst("id_mes")
    cDepartamento = rst("id_departamento")
    cProvincia = rst("id_provincia")
    cDistrito = rst("id_distrito")
    'cUrbanizacion = rst("id_urbanizacion")
    'cZona = rst("id_zona")
    Me.TxtDireccion1.Text = rst("direccion")
    
  
  
  
   
   
    
    StrPersonal = "si"
     Me.SstKardex.TabVisible(1) = True
  
  '-------------------------------------------
  
  
  
  If IsNull(rst("mail")) = False Then
      Me.txtEmail.Text = rst("mail")
    Else
    Me.txtEmail.Text = ""
  End If
  
If (rst("id_departamento") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lbldepartamento.Visible = False
    Me.DtcDepartamento.Visible = False
End If

If (rst("id_provincia") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lblprovincia.Visible = False
    Me.DtcProvincia.Visible = False
End If

If Trim(rst("id_credito")) = "si" Then
    If Len((rst("id_empresa_credito"))) = 11 Then
        Me.chkEmpresa.Value = True
        Me.txtRucEmpresa.Text = rst("id_empresa")
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRucEmpresa.Text) & "'"
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
        On Error GoTo nst
        Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
        img = Trim(rst("foto"))
    Else
nst:
        Me.Image1 = Nothing
    End If
End If
'--------- foto--------

If Val(cDepartamento) > 0 Then
    strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & cDepartamento & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDepartamento)
    Me.DtcDepartamento.BoundText = cDepartamento
End If

If Val(cProvincia) > 0 Then
    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE  id_departamento='" & cDepartamento & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProvincia)
    Me.DtcProvincia.BoundText = cProvincia
End If
If Val(cDistrito) > 0 Then
    strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE  id_provincia='" & cProvincia & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    Me.DtcDistrito.BoundText = cDistrito
End If





'Call CargarLogo(StrCodTabla)
Call LlenarTelefonos(Me.HfTelefonos, StrCodTabla)

  Exit Sub
'salir:   MsgBox "Se Presento un Problema Disculpe las molestias", vbInformation, KEY_EMPRESA
End Sub
Private Sub llenarFamiliares(ByVal Grilla As MSHFlexGrid)
strCadena = "SELECT F.id,P.nombre_completo,PR.descripcion as parentesco,F.telefono,F.dni_familia FROM persona_accidentes F,persona P,parentesco PR WHERE F.id_parentesco=PR.id_parentesco AND  F.dni='" & Me.txtRuc.Text & "' AND F.dni_familia=P.dni"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
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
        Next
        cabecera = "IDCODIGO" & vbTab & "DNI" & vbTab & "FAMILIAR" & vbTab & "PARENTESCO" & vbTab & "TELEFONO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id") & vbTab & rstT("dni_familia") & vbTab & rstT("nombre_completo") & vbTab & rstT("parentesco") & vbTab & rstT("telefono")
          Grilla.AddItem Fila
          Fila = ""
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
        

    Me.lblCumpleaños.Visible = True
    Me.txtdia.Visible = True
    Me.DtcMes.Visible = True
    Me.txtdia.Text = formato_item(rst("id_dia"), 2)
    Me.DtcMes.BoundText = formato_item(rst("id_mes"), 2)
    cDepartamento = rst("id_departamento")
    cProvincia = rst("id_provincia")
    cDistrito = rst("id_distrito")
    cUrbanizacion = Val(rst("id_urbanizacion"))
    cZona = Val(rst("id_zona"))
    Me.TxtDireccion1.Text = rst("direccion")
    Me.txtRuc.Text = rst("dni")
  
  If IsNull(rst("mail")) = False Then
      Me.txtEmail.Text = rst("mail")
    Else
    Me.txtEmail.Text = ""
  End If
  
If (rst("id_departamento") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lbldepartamento.Visible = False
    Me.DtcDepartamento.Visible = False
End If

If (rst("id_provincia") = 0) Or Me.LblCodPersona.Caption = "" Then
    Me.lblprovincia.Visible = False
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

Private Sub HfdPersona_Click()
If Me.HfdPersona.Rows > 0 Then
    Me.txtRucEmpresa.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
End If
End Sub



Private Sub OptSincredito_Click()
If Me.OptSincredito.Value = True Then
    Me.LblEmpresa.Visible = False
    Me.txtRucEmpresa.Visible = False
    Me.txtMaximoCredito.Visible = False
    
    
    
End If
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)

If Me.SSTab1.Tab = 1 Then
    strCadena = "SELECT Id as Codigo,CONCAT(nombre,' - ',Ejercicio) as Descripcion FROM con_periodo ORDER BY Ejercicio DESC,Mes DESC"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcPeriodoAsistencia)
End If

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
      Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtNDocumento_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcSexo.SetFocus
End If
End Sub

Private Sub TxtBusquedaRapida_Change()
trCadena = "SELECT E.cod_unico as Codigo,P.nombre_prod FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_EMPRESA & "' AND P.nombre_prod LIKE '%" & Trim(Me.TxtBusquedaRapida.Text) & "%' ORDER BY P.nombre_prod LIMIT 0,20 "
  Call ConfiguraRst(stracdena)
  Call LlenaDataCombo(Me.DtcEmpresaVinculada)
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Sub verificaTipo()

    StrPersonal = "si"
    strProveedor = "no"
    StrContable = "no"
    StrTransporte = "no"
    StrAlmacen = "no"
    StrAuspiciador = "no"
End Sub


Private Sub txtBuscar_Change()
 strCadena = "SELECT Per_Ruc,NombrePersona FROM Persona WHERE NombrePersona LIKE '%" & Trim(Me.txtBuscar.Text) & "%'ORDER BY NombrePersona ASC"
Call llenarGrid(Me.HfdPersona, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3050
        Next
         
         cabecera = "RAZON SOCIAL" & vbTab & "RUC"
        Grilla.AddItem cabecera
         For k = 0 To 1
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
              Fila = Fila & rst("Per_Ruc") & vbTab & rst("NombrePersona")
        
            If (Fila = "") Then
                X = 1
            End If
          Grilla.AddItem Fila
          
          Fila = ""
          rst.MoveNext
             
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
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

Private Sub txtdni_KeyPress(KeyAscii As Integer)
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
If Trim(Me.TxtDistrito.Text) <> "" Then
    strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE descripcion LIKE '%" & Trim(Me.TxtDistrito.Text) & "%'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    Set rst = Nothing
End If
End Sub

Private Sub TxtDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        Me.DtcDistrito.SetFocus
    End If
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
If frmpersonal.Procedencia <> modificar Then
strCadena = "SELECT dni FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Me.div_verifica.Visible = True
        Set rstT = Nothing
        strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.txtRuc.Text) & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Me.lblresultado.Caption = "Entidad ya forma parte de sus Clientes"
            Me.CmdVisualizar.Visible = True
        Else
            Me.lblresultado.Caption = "Registrado en www.Vitekey.com" + Chr(13)
            Me.CmdVisualizar.Visible = True
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
 If Len(Me.txtRuc.Text) = 8 Then
    Me.SstKardex.TabVisible(5) = True
Else
    Me.SstKardex.TabVisible(5) = False
 End If
 
 
 
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





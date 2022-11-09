VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmParteDiaria 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12620
      TabIndex        =   75
      Top             =   3480
      Width           =   255
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12360
      TabIndex        =   72
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox TxtDescripcionViaje 
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
      Left            =   8640
      MaxLength       =   80
      TabIndex        =   71
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox txtViajes 
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
      Left            =   7800
      MaxLength       =   80
      TabIndex        =   70
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox TxtMontoAlquiler 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   67
      Top             =   5480
      Width           =   1215
   End
   Begin VB.TextBox TxtCombustibleuso 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   65
      Top             =   4870
      Width           =   1215
   End
   Begin VB.TextBox TxtPrecioGasolina 
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
      Left            =   2400
      MaxLength       =   80
      TabIndex        =   62
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox TxtPrecioPetroleo 
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
      Left            =   2400
      MaxLength       =   80
      TabIndex        =   61
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox TxtCostoHora 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   59
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox TxtCombustibleHoy 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   57
      Top             =   4260
      Width           =   1215
   End
   Begin VB.TextBox TxtCapacidadMaquina 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   54
      Top             =   3630
      Width           =   1215
   End
   Begin VB.TextBox TxtCapacidadActual 
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
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   53
      Top             =   3945
      Width           =   1215
   End
   Begin VB.TextBox TxtInicio 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   49
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TxtFin 
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
      Left            =   12000
      MaxLength       =   80
      TabIndex        =   48
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Txtrecorrido 
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
      Left            =   9345
      MaxLength       =   80
      TabIndex        =   47
      Top             =   2550
      Width           =   1215
   End
   Begin VB.TextBox Txtcapacidad 
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
      Left            =   11985
      MaxLength       =   80
      TabIndex        =   46
      Top             =   2550
      Width           =   855
   End
   Begin VB.TextBox TxtHoras 
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
      Left            =   12960
      MaxLength       =   80
      TabIndex        =   45
      Top             =   1750
      Width           =   975
   End
   Begin VB.TextBox TxtHorometroFinal 
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
      Left            =   12000
      MaxLength       =   80
      TabIndex        =   6
      Top             =   1750
      Width           =   855
   End
   Begin VB.TextBox TxtId_Parte 
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
      Left            =   4560
      MaxLength       =   80
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox TxtGasolina 
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
      MaxLength       =   80
      TabIndex        =   10
      Top             =   4300
      Width           =   1095
   End
   Begin VB.TextBox TxtBusquedaPlaca 
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
      Left            =   12960
      MaxLength       =   80
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtGrasa 
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
      MaxLength       =   80
      TabIndex        =   12
      Top             =   4980
      Width           =   1095
   End
   Begin VB.TextBox TxtBusquedaSector 
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
      Left            =   6240
      MaxLength       =   80
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo DtcSector 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   3120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.TextBox TxtAceite 
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
      MaxLength       =   80
      TabIndex        =   11
      Top             =   4650
      Width           =   1095
   End
   Begin VB.TextBox TxtPetroleo 
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
      MaxLength       =   80
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtHorometro 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   5
      Top             =   1750
      Width           =   1215
   End
   Begin VB.TextBox TxtSerie 
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
      Left            =   4080
      MaxLength       =   80
      TabIndex        =   17
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TxtNumeroDoc 
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
      Left            =   4875
      MaxLength       =   80
      TabIndex        =   16
      Top             =   240
      Width           =   1050
   End
   Begin VB.TextBox TxtRucDestino 
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
      MaxLength       =   80
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtNombreDestino 
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
      MaxLength       =   80
      TabIndex        =   14
      Top             =   1635
      Width           =   5895
   End
   Begin VB.TextBox TxtDireccionDestino 
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
      MaxLength       =   80
      TabIndex        =   13
      Top             =   1950
      Width           =   5895
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   525
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5880
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   6000
      TabIndex        =   15
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121176065
      CurrentDate     =   41139
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12960
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":0000
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":08DA
            Key             =   "(GuiaRemision)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":0BF4
            Key             =   "(Imprimir)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   120
      TabIndex        =   18
      Top             =   6600
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   3270
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   780
         Left            =   120
         TabIndex        =   19
         Top             =   15
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1376
         ButtonWidth     =   1376
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Verificar"
               Key             =   "(Verificar)"
               ImageKey        =   "(GuiaRemision)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Grabar Ctrl+I"
               ImageKey        =   "(Imprimir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcAlmacenOrigen 
      Height          =   315
      Left            =   9360
      TabIndex        =   20
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSDataListLib.DataCombo DtcCantera 
      Height          =   315
      Left            =   9360
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5880
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":0C81
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":10D5
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":13F5
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":1849
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":1C9D
            Key             =   "(Atender)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":1FBD
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":22DD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteDiaria.frx":25FD
            Key             =   "(Declarar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   4065
      Left            =   13245
      TabIndex        =   21
      Top             =   3000
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7170
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4065
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   22
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Busqueda"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Anular)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcTransporte 
      Height          =   315
      Left            =   9360
      TabIndex        =   2
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   315
      Left            =   360
      TabIndex        =   43
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   1575
      Left            =   7800
      TabIndex        =   69
      Top             =   3840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "DESCRIPCION VIAJE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   8640
      TabIndex        =   74
      Top             =   3165
      Width           =   1485
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "VIAJES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   7800
      TabIndex        =   73
      Top             =   3165
      Width           =   510
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2460
      Left            =   7680
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MONTO ALQUILER :"
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
      Left            =   3660
      TabIndex        =   68
      Top             =   5535
      Width           =   2010
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   420
      Left            =   3480
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "COMBUSTIBLE USADO :"
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
      Left            =   3660
      TabIndex        =   66
      Top             =   4890
      Width           =   2010
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pre x Gl "
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
      Left            =   2910
      TabIndex        =   64
      Top             =   4320
      Width           =   600
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pre x Gl "
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
      Left            =   2910
      TabIndex        =   63
      Top             =   3960
      Width           =   600
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "COSTO X HORA :"
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
      Left            =   3660
      TabIndex        =   60
      Top             =   4620
      Width           =   2010
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CONBUSTIBLE + CARGA :"
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
      Left            =   3660
      TabIndex        =   58
      Top             =   4320
      Width           =   2010
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CAPACIDAD TANQUE :"
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
      Left            =   3660
      TabIndex        =   56
      Top             =   3720
      Width           =   2010
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "COMBUSTIBLE ACTUAL :"
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
      Left            =   3660
      TabIndex        =   55
      Top             =   4005
      Width           =   2010
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS MAQUINA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   3840
      TabIndex        =   52
      Top             =   3525
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   1740
      Left            =   3480
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "KILOMETRAJE :"
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
      Left            =   8040
      TabIndex        =   51
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "INICIO :"
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
      Left            =   8520
      TabIndex        =   50
      Top             =   2175
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "HOROM(FIN):"
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
      Left            =   10920
      TabIndex        =   44
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "GASOLINA :"
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
      Left            =   345
      TabIndex        =   41
      Top             =   4320
      Width           =   870
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "GRASA :"
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
      Left            =   600
      TabIndex        =   40
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ACEITE :"
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
      Left            =   570
      TabIndex        =   39
      Top             =   4695
      Width           =   645
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "COMBUSTIBLE Y LUBRICANTES"
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
      Left            =   360
      TabIndex        =   38
      Top             =   3645
      Width           =   2505
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PETROLEO :"
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
      Left            =   330
      TabIndex        =   37
      Top             =   3930
      Width           =   885
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "SECTOR :"
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
      Left            =   390
      TabIndex        =   36
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CAP. CARGA (M3):"
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
      Left            =   10590
      TabIndex        =   35
      Top             =   2550
      Width           =   1365
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "FIN :"
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
      Left            =   11595
      TabIndex        =   34
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "HOROM(INI) :"
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
      Left            =   8145
      TabIndex        =   33
      Top             =   1755
      Width           =   1020
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CANTERA :"
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
      Left            =   8355
      TabIndex        =   32
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS OPERADOR :"
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
      Left            =   480
      TabIndex        =   31
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "SUCURSAL        :"
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
      Left            =   7875
      TabIndex        =   30
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lblruc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DNI  :"
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
      Left            =   1020
      TabIndex        =   29
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label lblrazon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE :"
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
      Left            =   705
      TabIndex        =   28
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lbldireccion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Left            =   480
      TabIndex        =   27
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PLACA :"
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
      Left            =   8580
      TabIndex        =   26
      Top             =   1035
      Width           =   585
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "RECORRIDO :"
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
      Left            =   8130
      TabIndex        =   25
      Top             =   2550
      Width           =   1020
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION :"
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
      Left            =   2160
      TabIndex        =   24
      Top             =   6000
      Width           =   1185
   End
   Begin VB.Label lblanulado 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2100
      Left            =   120
      Top             =   780
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2100
      Left            =   7680
      Top             =   795
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   7680
      Top             =   120
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   7455
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   3540
      Left            =   120
      Top             =   3000
      Width           =   7455
   End
End
Attribute VB_Name = "FrmParteDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public cprod As String
Dim strMotivo As Integer









Private Sub DtcComprobanterel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub DtcComprobanteGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ' Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub CmdAgregar_Click()
If Val(Me.txtViajes.Text) > 0 And Trim(Me.TxtDescripcionViaje.Text) <> "" Then
   strCadena = "INSERT INTO parte_maquinaria_temporal(id_doc,serie,numero,viajes,descripcion,dni_save,ruc)VALUES('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Val(Me.txtViajes.Text) & "','" & Trim(Me.TxtDescripcionViaje.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
   CnBd.Execute (strCadena)
    
    
   Me.txtViajes.Text = 0
   Me.TxtDescripcionViaje.Text = ""
   Call Resalta(Me.txtViajes)
   Call llenarViajes(Me.HfdGrilla, 0)
End If
End Sub
Private Sub llenarViajes(ByVal Grilla As MSHFlexGrid, ByVal id_parte As Double)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

If id_parte > 0 Then
    strCadena = "SELECT id_detalle,viajes,descripcion FROM parte_maquinaria_detalle WHERE id_parte='" & id_parte & "' AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT id_temporal as id_detalle,viajes,descripcion FROM parte_maquinaria_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
End If
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
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 3800
           
           
          Next
         cabecera = "IDDETALLE" & vbTab & "Nº VIAJES" & vbTab & "DESCRIPCION"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = rst("id_detalle") & vbTab & rst("viajes") & vbTab & rst("descripcion")
             Grilla.AddItem Fila
             Fila = ""
             tTotal = tTotal + rst("viajes")
            rst.MoveNext
        Next i
             cabecera = "" & vbTab & tTotal & vbTab & "<---- VIAJES"
             Grilla.AddItem cabecera
             For k = 0 To 2
                 Grilla.col = k
                 Grilla.Row = i
                 Grilla.CellBackColor = &HDFDFE0
            Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Command1_Click()
If MsgBox("Esta Seguro de Eliminar este Item", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    strCadena = "DELETE FROM parte_maquinaria_temporal WHERE id_temporal='" & Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
     
    Call llenarViajes(Me.HfdGrilla, 0)
End If

End Sub

Private Sub DtcCantera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtHorometro)
End If
End Sub

Private Sub DtcSector_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtPetroleo)
End If
End Sub

Private Sub DtcTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM transporte WHERE id_transporte='" & Trim(Me.DtcTransporte.BoundText) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Me.TxtCapacidadMaquina.Text = Format(rstT("capacidad_tanque"), "###0.00")
        Me.TxtCostoHora.Text = Format(rstT("precio_hora"), "###0.00")
        strCadena = "SELECT combustible_actual FROM parte_maquinaria WHERE id_transporte='" & Trim(Me.DtcTransporte.BoundText) & "' AND ruc='" & KEY_RUC & "' ORDER BY id_parte DESC"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            rstT.MoveFirst
            Me.TxtCapacidadActual.Text = Format(rstT("combustible_actual"), "#,##0.00")
        Else
            Me.TxtCapacidadActual.Text = Format(0, "#,##0.00")
        End If
        
        If Me.DtcCantera.Enabled = True Then
            Me.DtcCantera.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.DTPicker1.Value = KEY_FECHA
strCadena = "SELECT * FROM combustible ORDER BY id_combustible ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtPrecioGasolina.Text = Format(rst("precio"), "#,##0.00")
    rst.MoveNext
    Me.TxtPrecioPetroleo.Text = Format(rst("precio"), "#,##0.00")
End If
strCadena = "SELECT T.id_transporte as Codigo,CONCAT(TT.descripcion,'-',T.placa) as Descripcion FROM transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTransporte)


strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND C.id_doc='0201' AND id_alm='" & KEY_ALM & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    MsgBox "CREE EL COMPROBANTE PARA ESTA SUCURSAL", vbInformation, KEY_EMPRESA
   
    Exit Sub
End If
Call LlenaDataCombo(Me.DtcTipoDoc)


strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0201' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
Me.TxtSerie.Text = rst("serie")
Me.TxtNumeroDoc.Text = rst("numero")
'Call Llenar_Temporal(Me.HfDetalle)

strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_proveedor='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCantera)

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacenOrigen)
Me.DtcAlmacenOrigen.BoundText = KEY_ALM
Me.DtcAlmacenOrigen.Locked = True
Me.TlbGrabar.Buttons("(Verificar)").Enabled = False

    strCadena = "SELECT id_urbanizacion as Codigo,descripcion as Descripcion FROM urbanizacion WHERE id_distrito='01754'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcSector)
    

End Sub

Private Sub BuscarResponsable(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = ruc
        FrmDetallePersona.ChkPersonal.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRucDestino.Text = rst("dni")
        Me.TxtNombreDestino.Text = rst("nombre_completo")
        Me.TxtDireccionDestino.Text = rst("direccion")
        Me.DtcTransporte.SetFocus
        Exit Sub
       
    End If

End Sub





Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
        Call nuevo
    Case KEY_UPDATE
         Procedencia = buscar
         FrmParteMaquinariaLista.Show
    Case KEY_DELETE
            Procedencia = anular
            FrmSeguridad.Show
            Exit Sub
    Case KEY_EXIT
        Unload Me
End Select
End Sub
Public Sub nuevo()
strCadena = "SELECT * FROM combustible ORDER BY id_combustible ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtPrecioGasolina.Text = Format(rst("precio"), "#,##0.00")
    rst.MoveNext
    Me.TxtPrecioPetroleo.Text = Format(rst("precio"), "#,##0.00")
End If
strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0201' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
Me.DtcTipoDoc.Enabled = True
Me.TxtSerie.Enabled = True
Me.TxtNumeroDoc.Enabled = True
Me.TxtSerie.Text = rst("serie")
Me.TxtNumeroDoc.Text = rst("numero")
Me.TxtRucDestino.Text = ""
Me.TxtNombreDestino.Text = ""
Me.TxtDireccionDestino.Text = ""
Me.TxtHorometro.Text = ""
Me.TxtHorometroFinal.Text = ""
Me.TxtInicio.Text = ""
Me.TxtCapacidadMaquina.Text = ""
Me.TxtCapacidadActual.Text = ""
Me.TxtCostoHora.Text = ""
Me.TxtCombustibleuso.Text = ""
Me.TxtCombustibleHoy.Text = ""
Me.TxtFin.Text = ""
Me.TxtMontoAlquiler.Text = ""
Me.Txtrecorrido.Text = ""
Me.Txtcapacidad.Text = ""
Me.TxtObservacion.Text = ""
Me.TxtAceite.Text = ""
Me.TxtGasolina.Text = ""
Me.TxtPetroleo.Text = ""
Me.TxtGrasa.Text = ""
Me.lblAnulado.Visible = False
Me.TxtId_parte.Text = ""
Call llenarViajes(Me.HfdGrilla, 0)
Me.DTPicker1.Value = KEY_FECHA
Call Resalta(Me.TxtRucDestino)
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False

End Sub
Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_SAVE
        Call Save
    
    Case "(Verificar)"
         Procedencia = Modificar
         FrmSeguridad.Show
         Exit Sub
         
    Case KEY_PRINT
      
        'Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text, Me.TxtNumeroDoc.Text, "00001")
        
End Select
End Sub
Private Sub Save()
Dim combustible As Single
If Me.DtcTipoDoc.BoundText = "" Or Me.TxtSerie.Text = "" Or Me.TxtNumeroDoc.Text = "" Then
   MsgBox "LLENE TODOS LOS PARAMETROS", vbInformation, KEY_EMPRESA
   Exit Sub
Else
    If Me.DtcAlmacenOrigen.BoundText = KEY_ALM Then
        combustible_actual = Val(Me.TxtCapacidadActual.Text) - Val(Me.TxtCombustibleuso.Text) + Val(Me.TxtCombustibleHoy.Text)
        strCadena = "INSERT INTO parte_maquinaria(fecha,id_doc,serie,numero,id_alm,id_transporte,id_operador,id_cantera,cantera,horometro_ini,horometro_fin,horas,precio,inicio,final,recorrido,capacidad,id_zona,petroleo,precio_petroleo,gasolina,pecio_gasolina,grasa,combustible_actual,uso_combustible,dni_save,observacion,ruc) " & _
        "VALUES('" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & KEY_ALM & "','" & Me.DtcTransporte.BoundText & "','" & Trim(Me.TxtRucDestino.Text) & "'," & _
        "'" & Trim(Me.DtcCantera.BoundText) & "','" & Trim(Me.DtcCantera.Text) & "','" & Val(Me.TxtHorometro.Text) & "','" & Val(Me.TxtHorometroFinal.Text) & "','" & Val(Me.TxtHoras.Text) & "','" & Val(Me.TxtCostoHora.Text) & "','" & Val(Me.TxtInicio.Text) & "','" & Val(Me.TxtFin.Text) & "','" & Val(Me.Txtrecorrido.Text) & "','" & Val(Me.Txtcapacidad.Text) & "','" & Trim(Me.DtcSector.BoundText) & "','" & Val(Me.TxtPetroleo.Text) & "','" & Val(Me.TxtPrecioPetroleo.Text) & "','" & Val(Me.TxtGasolina.Text) & "','" & Val(Me.TxtPrecioGasolina.Text) & "','" & Val(Me.TxtGrasa.Text) & "','" & combustible_actual & "','" & Val(Me.TxtCombustibleuso.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.TxtObservacion.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        
        Me.TxtId_parte.Text = LastRegistro("parte_maquinaria", "id_parte")
        strCadena = "SELECT * FROM parte_maquinaria_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO parte_maquinaria_detalle(id_parte,viajes,descripcion,ruc)VALUES('" & Val(Me.TxtId_parte.Text) & "','" & rst("viajes") & "','" & rst("descripcion") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                 
                rst.MoveNext
            Next i
            strCadena = "DELETE FROM parte_maquinaria_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
             
             
        End If
    Else
         'strCadena = "UPDATE parte_maquinaria SET finalizado='si',id_recibio='" & KEY_USUARIO & "' WHERE id_transferencia='" & Val(Me.TxtId_transferencia.text) & "' ANd ruc='" & KEY_RUC & "'"
         'CnBd.Execute (strCadena)
         
         'Call savedetalle(Val(Me.TxtId_transferencia.text))
         Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
         Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
         
         Exit Sub
    End If
    
    
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacenOrigen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
     
     
    'Call savedetalle(id_transferencia)
End If

End Sub
Private Sub savedetalle(ByVal id_transferencia As Double)
strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO movimiento_transferencia_detalle(id_transferencia,id_producto,cantidad,recibido,peso,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("recibido") & "','" & rstT("peso") & "','" & rstT("total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
            
            
           rstT.MoveNext
        Next i
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
         
    End If
End Sub


Private Sub TxtBusquedaPlaca_Change()
strCadena = "SELECT T.id_transporte as Codigo,CONCAT(TT.descripcion,'-',T.placa) as Descripcion FROM transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "' AND T.placa LIKE '%" & Trim(Me.TxtBusquedaPlaca.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTransporte)

End Sub

Private Sub TxtBusquedaSector_Change()
 strCadena = "SELECT id_urbanizacion as Codigo,descripcion as Descripcion FROM urbanizacion WHERE descripcion LIKE '%" & Trim(Me.TxtBusquedaSector.Text) & "%'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcSector)
End Sub

Private Sub TxtFin_Change()
Me.Txtrecorrido.Text = Format(Val(Me.TxtFin.Text) - Val(Me.TxtInicio.Text), "###0.00")
End Sub

Private Sub TxtGasolina_Change()
Me.TxtCombustibleHoy.Text = Format(Val(Me.TxtCapacidadActual.Text) + Val(Me.TxtGasolina.Text) / Val(Me.TxtPrecioGasolina.Text) + Val(Me.TxtPetroleo.Text) / Val(Me.TxtPrecioPetroleo.Text), "###0.00")
End Sub

Private Sub TxtHorometro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtHorometroFinal)
End If
End Sub

Private Sub TxtHorometroFinal_Change()
Me.TxtHoras.Text = Format(Val(Me.TxtHorometroFinal.Text) - Val(Me.TxtHorometro.Text), "#,##0.00")
Me.TxtMontoAlquiler.Text = Format(Val(Me.TxtHoras.Text) * Val(Me.TxtCostoHora.Text), "#,##0.00")
strCadena = "SELECT * FROM transporte WHERE id_transporte='" & Trim(Me.DtcTransporte.BoundText) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.TxtCombustibleuso.Text = Format(rstT("consumo_hora") * Val(Me.TxtHoras.Text), "#,##0.00")
End If

End Sub

Private Sub TxtHorometroFinal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtInicio)
End If
End Sub

Private Sub TxtKilometraje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcSector.SetFocus
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_comprobante
End If
End Sub
Public Sub buscar_comprobante(Optional id_transferencia As Double)
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
    strCadena = "SELECT * FROM parte_maquinaria WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    
        Me.TxtId_parte.Text = rst("id_parte")
        Me.DtcTipoDoc.BoundText = rst("id_doc")
        Me.DTPicker1.Value = rst("fecha")
        Me.TxtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        
        If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
            Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
        Else
            Me.lblAnulado.Visible = False
            Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
        End If
        
        
        
        
        Me.TxtObservacion.Text = rst("observacion")
        If IsNull(rst("id_operador")) = False And rst("id_operador") <> "" Then
            Me.TxtRucDestino.Text = rst("id_operador")
            Me.TxtNombreDestino.Text = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_operador"))
            Me.TxtDireccionDestino.Text = BDBuscarCampo("persona", "direccion", "dni", rst("id_operador"))
            
        Else
            Me.TxtNombreDestino.Text = ""
        End If
        
        If IsNull(rst("id_transporte")) = False And rst("id_transporte") <> "" Then
            Me.DtcTransporte.BoundText = rst("id_transporte")
            Me.DtcCantera.BoundText = rst("id_cantera")
            Me.TxtHorometro.Text = rst("horometro_ini")
            Me.TxtHorometroFinal.Text = rst("horometro_fin")
            Me.TxtInicio.Text = rst("inicio")
            Me.TxtFin.Text = rst("final")
            Me.TxtCombustibleuso.Text = Format(rst("uso_combustible"), "#,##0.00")
            Me.TxtCapacidadMaquina.Text = Format(BDBuscarCampoRuc("transporte", "capacidad_tanque", "id_transporte", rst("id_transporte")), "#,##0.00")
            Me.TxtCapacidadActual.Text = Format(rst("combustible_actual"), "#,##0.00")
            Me.TxtCostoHora.Text = Format(BDBuscarCampoRuc("transporte", "precio_hora", "id_transporte", rst("id_transporte")), "#,##0.00")
            Me.TxtMontoAlquiler.Text = Format(rst("horas") * rst("precio"), "#,##0.00")
            Me.Txtrecorrido.Text = rst("recorrido")
            Me.Txtcapacidad.Text = rst("capacidad")
            Me.DtcSector.BoundText = rst("id_zona")
            Me.TxtPetroleo.Text = rst("petroleo")
            Me.TxtGasolina.Text = rst("gasolina")
            Me.TxtAceite.Text = rst("aceite")
            Me.TxtGrasa.Text = rst("grasa")
            Me.TxtObservacion.Text = rst("observacion")
        End If
        Call llenarViajes(Me.HfdGrilla, rst("id_parte"))
        Me.DtcTipoDoc.Enabled = False
        Me.TxtSerie.Enabled = False
        Me.TxtNumeroDoc.Enabled = False
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
        Set rst = Nothing
    'Else
   ' Call Resalta(Me.TxtRuc)
    'End If
    End If
End Sub




Private Sub TxtPetroleo_Change()
Me.TxtCombustibleHoy.Text = Format(Val(Me.TxtCapacidadActual.Text) + Val(Me.TxtGasolina.Text) / Val(Me.TxtPrecioGasolina.Text) + Val(Me.TxtPetroleo.Text) / Val(Me.TxtPrecioPetroleo.Text), "###0.00")

End Sub

Private Sub TxtPetroleo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
End If
End Sub

Private Sub TxtRucDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarResponsable(Trim(Me.TxtRucDestino.Text))
End If
End Sub




Public Sub Llenar_Temporal(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM movimiento_transferencia_temporal T,producto P,unidad U WHERE T.id_producto=P.id_producto AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND T.id_doc='" & Me.DtcTipoDoc.BoundText & "' AND T.serie='" & Me.TxtSerie.Text & "' AND T.numero='" & Me.TxtNumeroDoc.Text & "' AND T.dni_save='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1400
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "ENVIADO" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "PESO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("peso") * rst("cantidad")
          Fila = rst("id_temporal") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("recibido"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & Format(rst("peso") * rst("cantidad"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("recibido") Then
          For k = 0 To 7
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
      For k = 6 To 7
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0C0FF
      Next k
      Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenar_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_transferencia As Double)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM movimiento_transferencia_detalle T,producto P,unidad U WHERE T.id_producto=P.id_producto AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND T.id_transferencia='" & id_transferencia & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1400
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "ENVIADO" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "PESO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("total")
          Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("recibido"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("recibido") Then
          For k = 0 To 7
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
     
      'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
      'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.TxtNumeroDoc)
End If
End Sub



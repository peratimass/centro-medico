VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmTransacciones 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "GIRADO A :"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   11775
      Begin VB.Label lblRazonSocial 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2640
         TabIndex        =   24
         Top             =   600
         Width           =   7545
      End
      Begin VB.Label lblRuc 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE / RAZON SOCIAL :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   600
         Width           =   2010
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "RUC/DNI :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         TabIndex        =   21
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECCIONE CUENTA BANCARIA"
      ForeColor       =   &H00800000&
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   11775
      Begin VB.Frame Frame4 
         Caption         =   "CUENTA DESTINO (BENEFICIARIO)   *"
         ForeColor       =   &H00800000&
         Height          =   2415
         Left            =   360
         TabIndex        =   36
         Top             =   2760
         Width           =   4695
         Begin VB.TextBox TxtCuentaBancaria 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   38
            Top             =   2010
            Width           =   2415
         End
         Begin VB.CommandButton cmdagregar 
            Caption         =   "+"
            Height          =   320
            Left            =   4200
            TabIndex        =   37
            Top             =   2040
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshCuentasBancarias 
            Height          =   975
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   1720
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
         Begin MSDataListLib.DataCombo DtcBanco 
            Height          =   315
            Left            =   1680
            TabIndex        =   40
            Top             =   1280
            Width           =   2415
            _ExtentX        =   4260
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
            Left            =   1680
            TabIndex        =   43
            Top             =   1630
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            ListField       =   ""
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
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONEDA :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   825
            TabIndex        =   44
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CUENTA BANCARIA:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   75
            TabIndex        =   42
            Top             =   2100
            Width           =   1515
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANCO   :"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   855
            TabIndex        =   41
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.TextBox TxtComision 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3120
         TabIndex        =   33
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox TxtCentroCostos 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6720
         TabIndex        =   26
         Top             =   720
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "      TRANSFERENCIA A OTRA CUENTA"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   6720
         TabIndex        =   18
         Top             =   3240
         Width           =   4695
         Begin VB.CheckBox ChkPasar 
            Appearance      =   0  'Flat
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   0
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DtcMisCuentas 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
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
      End
      Begin VB.TextBox TxtObservacion 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   6840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   4200
         Width           =   4575
      End
      Begin VB.TextBox TxtItf 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox TxtDepositado 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3120
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox TxtMovimiento 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3120
         TabIndex        =   13
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TxtTc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3120
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox TxtOperacion 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6720
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo DtcCuentaBancaria 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   360
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
      Begin MSComctlLib.ImageList ImgIconos 
         Left            =   10560
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":0000
               Key             =   "(Aceptar)"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":031C
               Key             =   "(Eliminar)"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":077C
               Key             =   "(Imprimir)"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":0809
               Key             =   "(Inicio)"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":0C69
               Key             =   "(Modificar)"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":0F85
               Key             =   "(Nuevo)"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":13E5
               Key             =   "(Quitar)"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":1701
               Key             =   "(Salir)"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":1B61
               Key             =   "(Red)"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":1FC1
               Key             =   "(Grabar)"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":28A1
               Key             =   "(Agregar)"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":2BBD
               Key             =   "(Buscar)"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmTransacciones.frx":2ED9
               Key             =   "(Cancelar)"
            EndProperty
         EndProperty
      End
      Begin ComCtl3.CoolBar ClbAcciones 
         Height          =   870
         Left            =   8880
         TabIndex        =   29
         Top             =   4920
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1535
         BandCount       =   1
         ForeColor       =   -2147483635
         ImageList       =   "ImgIconos"
         FixedOrder      =   -1  'True
         VariantHeight   =   0   'False
         EmbossPicture   =   -1  'True
         _CBWidth        =   2475
         _CBHeight       =   870
         _Version        =   "6.0.8169"
         Child1          =   "TlbGrabar"
         MinHeight1      =   810
         Width1          =   3180
         FixedBackground1=   0   'False
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar TlbGrabar 
            Height          =   810
            Left            =   30
            TabIndex        =   30
            Top             =   30
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   1429
            ButtonWidth     =   1402
            ButtonHeight    =   1429
            Style           =   1
            ImageList       =   "ImgIconos"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
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
                  Caption         =   "Imprimir"
                  Key             =   "(Imprimir)"
                  Object.ToolTipText     =   "Cancelar"
                  ImageKey        =   "(Imprimir)"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Salir"
                  Key             =   "(Salir)"
                  ImageKey        =   "(Salir)"
               EndProperty
            EndProperty
            OLEDropMode     =   1
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
         Height          =   1575
         Left            =   6720
         TabIndex        =   35
         Top             =   1560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2778
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
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(*) Seleccionar la cuenta Destino"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   5520
         Width           =   4695
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "COMISION :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1920
         TabIndex        =   34
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label lblcostos 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   6720
         TabIndex        =   28
         Top             =   1080
         Width           =   4905
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "C.COSTOS  :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5610
         TabIndex        =   27
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "OBSERVACION:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5520
         TabIndex        =   17
         Top             =   4440
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ITF :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2460
         TabIndex        =   11
         Top             =   1920
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL DEPOSITADO :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL MOVIMIENTO :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TIPO CAMBIO :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1695
         TabIndex        =   8
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº OPERACION :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5310
         TabIndex        =   7
         Top             =   360
         Width           =   1230
      End
   End
   Begin MSComCtl2.DTPicker DtpOperacion 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Format          =   169607169
      CurrentDate     =   41132
   End
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Format          =   169607169
      CurrentDate     =   41132
   End
   Begin VB.Label lblcheque 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   9600
      TabIndex        =   32
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nº CHEQUE:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8640
      TabIndex        =   31
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "FECHA VALOR :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "FECHA OPERACION :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "FrmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChkPasar_Click()
If Me.DtcMisCuentas.Enabled = False Then
    Me.DtcMisCuentas.Enabled = True
Else
    Me.DtcMisCuentas.Enabled = False
End If
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal id_cheque As Double)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM cheque_detalle WHERE id_cheque='" & id_cheque & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
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
           Grilla.ColWidth(1) = 500
           Grilla.ColWidth(2) = 3200
           Grilla.ColWidth(3) = 1100
       Next
         cabecera = "IDITEM" & vbTab & "ITEM" & vbTab & "DESCRIPCION" & vbTab & "MONTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
        tTotal = tTotal + rst("monto")
             Fila = rst("id_item") & vbTab & formato_item(i, 2) & vbTab & rst("detalle") & vbTab & Format(rst("monto"), "#,##0.00")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL A GIRAR  ***********" & vbTab & Format(tTotal, "###0.00")
       Grilla.AddItem Fila
       For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
      Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdAgregar_Click()
If Len(Me.TxtCuentaBancaria.Text) > 5 Then
    strCadena = "SELECT * FROM persona_cuentabancaria WHERE dni='" & Trim(Me.lblruc.Caption) & "' AND cuenta='" & Trim(Me.TxtCuentaBancaria.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO persona_cuentabancaria(dni,id_banco,id_moneda,cuenta)VALUES('" & Trim(Me.lblruc.Caption) & "','" & Me.DtcBanco.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Trim(Me.TxtCuentaBancaria.Text) & "') "
        CnBd.Execute (strCadena)
         
        Call llenar_cuentas(Me.MshCuentasBancarias, Trim(Me.lblruc.Caption))
    End If
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1500
Dim residuo As Single, itf As Single
Me.lblcheque.Caption = FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 2)
strCadena = "SELECT * FROM cheque C,persona P WHERE C.id_cheque='" & Val(FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND C.id_beneficiario=P.dni"
Call ConfiguraRstT(strCadena)
Me.lblruc.Caption = rstT("dni")
Me.LblRazonSocial.Caption = UCase(rstT("nombre_completo"))
Me.TxtMovimiento.Text = Format(rstT("monto"), "###0.00")
Me.TxtDepositado.Text = Format(rstT("saldo"), "###0.00")
Me.TxtCentroCostos.Text = rstT("ccostos")
Me.lblcostos.Caption = BDBuscarCampo("plan_contable_det", "plan_des", "pc_codigo", rstT("ccostos"))

residuo = Val(Me.TxtDepositado.Text) Mod 1000
If ((Val(Me.TxtDepositado.Text) - residuo) > 0) Then
    itf = (Val(Me.TxtDepositado.Text) - residuo) * 0.005 / 100
Else
    itf = 0
End If
  
Me.TxtItf.Text = Format(itf, "#,##0.000")
strCadena = "SELECT M.id_cuenta as Codigo,CONCAT(M.descripcion,'-',MO.descripcion,'-',M.numero_cuenta) as Descripcion FROM mis_cuentas M,moneda MO,cheque C,chequera CH WHERE C.id_chequera=CH.id_chequera AND CH.id_cuenta=M.id_cuenta AND C.id_cheque='" & FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0) & "' AND CH.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND  M.id_moneda=MO.id_moneda  AND M.id_tipo<>'01'AND M.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentaBancaria)

Me.txtTc.Text = Format(KEY_CAMBIO, "#,##0.000")
strCadena = "SELECT M.id_cuenta as Codigo,CONCAT(M.descripcion,'-',MO.descripcion,' ',M.numero_cuenta) as Descripcion FROM mis_cuentas M,moneda MO WHERE  M.id_moneda=MO.id_moneda  AND M.ruc='" & KEY_RUC & "' AND M.id_cuenta<>'" & Me.DtcCuentaBancaria.BoundText & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMisCuentas)
Me.Caption = KEY_EMPRESA
Call llenarGrid(Me.HfDetalle, Val(FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0)))
Me.DtcMisCuentas.Enabled = False
Call llenar_cuentas(Me.MshCuentasBancarias, Trim(Me.lblruc.Caption))
strCadena = "SELECT id_banco as Codigo,abreviatura as Descripcion FROM banco ORDER BY abreviatura"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcBanco)
strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
End Sub
Public Sub llenar_cuentas(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir
strCadena = "SELECT B.id_banco,B.abreviatura,M.descripcion,PB.cuenta,M.id_moneda FROM persona_cuentabancaria PB,banco B,moneda M WHERE PB.id_banco=B.id_banco AND PB.id_moneda=M.id_moneda AND PB.dni='" & dni & "'"
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
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 1000
            Grilla.ColWidth(3) = 2000
            Grilla.ColWidth(4) = 0
        Next
        cabecera = "BCO" & vbTab & "MONEDA" & vbTab & "CUENTA"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_banco") & vbTab & rst("abreviatura") & vbTab & rst("descripcion") & vbTab & rst("cuenta") & vbTab & rst("id_moneda")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
      
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
Private Sub Save()
Dim Documento As String
Dim abreviatura As String, cuenta_destino As String, operacion As String, monto_pago As Double, Saldo As Double, id_moneda As String

If Val(Me.TxtDepositado.Text) > 0 Then
    monto_pago = Val(Me.TxtDepositado.Text)
    'strCadena = "SELECT * FROM almacen_comprobante A,comprobantes C WHERE A.id_doc='0096' AND A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "'"
    'Call ConfiguraRst(strCadena)
    cuenta_destino = Me.DtcBanco.Text + "-" + Me.DtcMoneda.Text + "-" + Me.TxtCuentaBancaria.Text
    
     Documento = "CH :" + Space(1) + Me.lblcheque.Caption
     operacion = formato_item(Me.txtOperacion.Text, 5)
    
    
    
    '************** CANCELAR FACTURAS ******************
    strCadena = "SELECT M.id_compra,M.saldo,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.id_moneda FROM movimiento_compra M,cheque_factura CH,comprobantes C WHERE M.id_doc=C.id_doc AND  M.id_compra=CH.id_compra AND M.ruc='" & KEY_RUC & "' AND CH.ruc='" & KEY_RUC & "' AND id_cheque='" & FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0) & "' ORDER BY M.saldo ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
         For i = 0 To rst.RecordCount - 1
                '------ VERIFICAR MONEDA
                id_moneda = Trim(BDBuscarCampo("mis_cuentas", "id_moneda", "id_cuenta", Me.DtcMisCuentas.BoundText))
                If rst("id_moneda") <> id_moneda Then
                        If rst("id_moneda") = "00001" Then
                               Saldo = rst("saldo") / KEY_CAMBIO
                        Else
                               Saldo = rst("saldo") * KEY_CAMBIO
                        End If
                  Else
                    Saldo = rst("saldo")
                End If
                '-------END
                If monto_pago > 0 Then
                        If monto_pago >= Saldo Then
                            saldof = 0
                            monto_pagado = Saldo
                            monto_pago = monto_pago - Saldo
                        Else
                            saldof = Saldo - monto_pago
                            monto_pagado = monto_pago
                            monto_pago = 0
                        End If
                    End If
            strCadena = "INSERT INTO mis_cuentas_det(documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc," & _
            "operacion,id_movimiento,dni_save,ccostos,cuenta_destino,ruc)VALUES('" & rst("comprobante") & "'," & _
            "'" & Me.DtcCuentaBancaria.BoundText & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "','" & KEY_FECHA & "','" & Me.lblruc.Caption & "'," & _
            "'" & Trim(Me.txtObservacion.Text) & "','" & monto_pagado & "','" & monto_pagado * -1 & "','" & Val(Me.txtTc.Text) & "'" & _
            ",'" & operacion & "','" & rst("id_compra") & "','" & KEY_USUARIO & "','" & Trim(Me.TxtCentroCostos.Text) & "','" & cuenta_destino & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
                    
                    If rst("id_moneda") <> id_moneda Then
                        If rst("id_moneda") = "00001" Then
                               saldof = saldof * KEY_CAMBIO
                        Else
                            saldof = saldof / KEY_CAMBIO
                        End If
                    End If
                    strCadena = "UPDATE movimiento_compra SET saldo='" & saldof & "' WHERE id_compra='" & rst("id_compra") & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                     
            rst.MoveNext
        Next i
    Else
            strCadena = "INSERT INTO mis_cuentas_det(documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc," & _
            "operacion,id_movimiento,dni_save,ccostos,cuenta_destino,ruc)VALUES('" & Documento & "'," & _
            "'" & Me.DtcCuentaBancaria.BoundText & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "','" & KEY_FECHA & "','" & Me.lblruc.Caption & "'," & _
            "'" & Trim(Me.txtObservacion.Text) & "','" & Val(Me.TxtDepositado.Text) & "','" & Val(Me.TxtDepositado.Text) * -1 & "','" & Val(Me.txtTc.Text) & "'" & _
            ",'" & operacion & "','" & rst("id_compra") & "','" & KEY_USUARIO & "','" & Trim(Me.TxtCentroCostos.Text) & "','" & cuenta_destino & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
    End If
    '************** FIN
    
    
    If Val(Me.TxtItf.Text) > 0 Then
        operacion = formato_item(Val(operacion) + 1, 5)
        Documento = "ITF" + Space(1) + Documento
        strCadena = "INSERT INTO mis_cuentas_det(documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc," & _
    "operacion,dni_save,ccostos,ruc)VALUES('" & Documento & "'," & _
    "'" & Me.DtcCuentaBancaria.BoundText & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "','" & KEY_FECHA & "','" & Me.lblruc.Caption & "'," & _
    "'" & Documento & "','" & Val(Me.TxtItf.Text) & "','" & Val(Me.TxtItf.Text) * -1 & "','" & Val(Me.txtTc.Text) & "'" & _
    ",'" & operacion & "','" & KEY_USUARIO & "','64121','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    End If
    If Val(Me.TxtComision.Text) > 0 Then
        operacion = formato_item(Val(operacion) + 1, 5)
        Documento = "COMISION" + Space(1) + Documento
        strCadena = "INSERT INTO mis_cuentas_det(documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc," & _
    "operacion,dni_save,ccostos,ruc)VALUES('" & Documento & "'," & _
    "'" & Me.DtcCuentaBancaria.BoundText & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "','" & KEY_FECHA & "','" & Me.lblruc.Caption & "'," & _
    "'" & Documento & "','" & Val(Me.TxtComision.Text) & "','" & Val(Me.TxtComision.Text) * -1 & "','" & Val(Me.txtTc.Text) & "'" & _
    ",'" & operacion & "','" & KEY_USUARIO & "','64121','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    End If
    
    
    strCadena = "UPDATE cheque SET saldo='0' WHERE id_cheque='" & Val(FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    End If

End Sub

Private Sub MshCuentasBancarias_SelChange()

If Trim(Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 2) <> "") Then
    Me.DtcBanco.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 0)
    Me.DtcMoneda.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 4)
    Me.TxtCuentaBancaria.Text = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 3)
    Call Resalta(Me.txtOperacion)
Else
    Me.TxtCuentaBancaria.Text = ""
End If
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_SAVE
         Call Save
    Case KEY_EXIT
        Unload Me
End Select
End Sub

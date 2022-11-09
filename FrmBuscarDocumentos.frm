VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBuscarDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Documentos"
   ClientHeight    =   3060
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   7425
   Icon            =   "FrmBuscarDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7425
   Begin TabDlg.SSTab SstBusca 
      Height          =   2655
      Left            =   168
      TabIndex        =   0
      Top             =   203
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmBuscarDocumentos.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ChkEstado"
      Tab(0).Control(1)=   "DtcEstado"
      Tab(0).Control(2)=   "TxtNDocumento"
      Tab(0).Control(3)=   "TxtEntidad"
      Tab(0).Control(4)=   "LblNDocumento"
      Tab(0).Control(5)=   "LblEntidad"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Total Documento"
      TabPicture(1)   =   "FrmBuscarDocumentos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtInferior"
      Tab(1).Control(1)=   "TxtSuperior"
      Tab(1).Control(2)=   "OptInferior"
      Tab(1).Control(3)=   "OptSuperior"
      Tab(1).Control(4)=   "OptIntervaloMonto"
      Tab(1).Control(5)=   "TxtMaximo"
      Tab(1).Control(6)=   "TxtMinimo"
      Tab(1).Control(7)=   "LblMaximo"
      Tab(1).Control(8)=   "LblMinimo"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Fechas"
      TabPicture(2)   =   "FrmBuscarDocumentos.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "LblDesde"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LblHasta"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DtpDesde"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DtpHasta"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ChkFecha"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CheckBox ChkEstado 
         Caption         =   "Estado :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DtcEstado 
         Height          =   315
         Left            =   -73200
         TabIndex        =   17
         Top             =   1800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ForeColor       =   12582912
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox TxtInferior 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73440
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox TxtSuperior 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73440
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton OptInferior 
         Caption         =   "Inferior a:                S/."
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   -74640
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton OptSuperior 
         Caption         =   "Superior a:                S/."
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   -74640
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton OptIntervaloMonto 
         Caption         =   "Intervalo de PagoTotal"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtMaximo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -70800
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtMinimo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73440
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtNDocumento 
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   -73200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox ChkFecha 
         BackColor       =   &H80000004&
         Caption         =   "Por Rango de Fechas de Emisión"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox TxtEntidad 
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   -73200
         MaxLength       =   80
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   3600
         TabIndex        =   19
         Top             =   1035
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   63176705
         CurrentDate     =   37327
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   1035
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   63176705
         CurrentDate     =   37327
      End
      Begin VB.Label LblMaximo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Hasta :   S/."
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -71760
         TabIndex        =   9
         Top             =   855
         Width           =   870
      End
      Begin VB.Label LblMinimo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Desde :  S/."
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74400
         TabIndex        =   8
         Top             =   840
         Width           =   870
      End
      Begin VB.Label LblNDocumento 
         AutoSize        =   -1  'True
         Caption         =   "N° de Documento :"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   7
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label LblHasta 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Hasta : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3000
         TabIndex        =   5
         Top             =   1095
         Width           =   555
      End
      Begin VB.Label LblDesde 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Desde : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   1095
         Width           =   600
      End
      Begin VB.Label LblEntidad 
         AutoSize        =   -1  'True
         Caption         =   "Entidad:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74760
         TabIndex        =   2
         Top             =   840
         Width           =   585
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6202
      Top             =   1763
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
            Picture         =   "FrmBuscarDocumentos.frx":035E
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":07B2
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":0AD2
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":0F26
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":137A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":169A
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":19BA
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":1CDA
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarDocumentos.frx":1FFA
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1890
      Left            =   6356
      TabIndex        =   21
      Top             =   585
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   3334
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   1890
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2340
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   4128
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  Buscar   "
               Key             =   "(Buscar)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "(Cancelar)"
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
End
Attribute VB_Name = "FrmBuscarDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Estado As String, DteInicio As Date, DteFin As Date
Public MontoMinimo As Double, MontoMaximo As Double

Private Sub ChkEstado_Click()
  If ChkEstado.Value = 1 Then
    DtcEstado.Enabled = True
  Else
    DtcEstado.Enabled = False
  End If
End Sub

Private Sub ChkFecha_Click()
  If ChkFecha.Enabled = True Then
    DtpDesde.Enabled = True
    DtpHasta.Enabled = True
  Else
    DtpDesde.Enabled = False
    DtpHasta.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  StrCadena = "SELECT cestado as Codigo, sdescripcion as Descripcion FROM Estado " & _
  " WHERE  NOT cestado LIKE '%L%' ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcEstado)
  DtcEstado.Enabled = False
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  Select Case MDIFrmPrincipal.EnumBuscar
    Case BDocumentoCompra
      LblEntidad.Caption = "Distribuidora"
      SstBusca.TabVisible(1) = True
    Case BDocumentoVenta
      LblEntidad.Caption = "Cliente"
      SstBusca.TabVisible(1) = True
    Case BOtraEntrada
      SstBusca.TabVisible(1) = False
    Case BOtraSalida
      SstBusca.TabVisible(1) = False
    Case BPedido
      LblEntidad.Caption = "Cliente"
      SstBusca.TabVisible(1) = True
  End Select
  CenterForm Me
End Sub

Private Sub OptInferior_Click()
  TxtInferior.Enabled = True
  TxtSuperior.Enabled = False
  TxtMaximo.Enabled = False
  TxtMinimo.Enabled = False
End Sub

Private Sub OptIntervaloMonto_Click()
  TxtInferior.Enabled = False
  TxtSuperior.Enabled = False
  TxtMaximo.Enabled = True
  TxtMinimo.Enabled = True
End Sub

Private Sub OptSuperior_Click()
  TxtInferior.Enabled = False
  TxtSuperior.Enabled = True
  TxtMaximo.Enabled = False
  TxtMinimo.Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case KEY_BROWSER
    Estado = ""
    DteInicio = DTEMINIMA
    DteFin = DTEMAXIMA
    MontoMinimo = 0
    MontoMaximo = 1000000
    If ChkEstado.Value = 1 Then
      Estado = Replace(DtcEstado.BoundText, "'", "''")
    End If
    If ChkFecha.Value = 1 Then
       DteInicio = DtpDesde.Value
       DteFin = DtpHasta.Value
    End If
    If OptInferior.Value = True Then
      If Not TxtInferior.Text = "" Then
        MontoMaximo = TxtInferior.Text
      End If
    End If
    If OptIntervaloMonto.Value = True Then
      If Not (TxtMinimo.Text = "" Or TxtMaximo.Text = "") Then
        MontoMinimo = TxtMinimo.Text
        MontoMaximo = TxtMaximo.Text
      End If
    End If
    If OptSuperior.Value = True Then
      If Not TxtSuperior.Text = "" Then
        MontoMinimo = TxtSuperior.Text
      End If
    End If
    Me.Hide
    Select Case MDIFrmPrincipal.EnumBuscar
      Case BDocumentoCompra
        FrmDocumentoCompra.Show
      Case BDocumentoVenta
        FrmDocumentoVenta.Show
      Case BOtraEntrada
        'FrmOtraEntrada.Show
      Case BOtraSalida
        'FrmOtraSalida.Show
      Case BPedido
        FrmPedido.Show
    End Select
  Case KEY_CANCEL
    Unload Me
  End Select
End Sub

Private Sub TxtInferior_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtMaximo__KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtMinimo__KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtNDocumento__KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtSuperior__KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtNDocumento_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

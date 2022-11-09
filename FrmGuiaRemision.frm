VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleGuia 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATOS CHOFER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   6960
      TabIndex        =   33
      Top             =   6720
      Width           =   6135
      Begin VB.TextBox TxtNombreChofer 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         MaxLength       =   100
         TabIndex        =   43
         Top             =   840
         Width           =   5445
      End
      Begin VB.TextBox TxtLicencia 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   42
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licencia:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2805
         TabIndex        =   44
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EMPRESA DE TRAMPORTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   360
      TabIndex        =   25
      Top             =   5160
      Width           =   12735
      Begin VB.TextBox TxtDireccionTransporte 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   31
         Top             =   720
         Width           =   4725
      End
      Begin VB.TextBox TxtRuc_Transportes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8085
         MaxLength       =   11
         TabIndex        =   29
         Top             =   360
         Width           =   1965
      End
      Begin VB.TextBox TxtNombreEmpresaTransporte 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   27
         Top             =   360
         Width           =   4725
      End
      Begin VB.TextBox TxtCodigoEmpresaTransporte 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   26
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   375
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7200
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMPROBANTE REFERENCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   8640
      TabIndex        =   21
      Top             =   3480
      Width           =   4455
      Begin VB.TextBox TxtSerie_Factura 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   24
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox TxtNumero_Factura 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2265
         MaxLength       =   10
         TabIndex        =   23
         Top             =   360
         Width           =   1485
      End
      Begin MSDataListLib.DataCombo DtcTipoDoc_Ref 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DtpFecha_Factura 
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   175964161
         CurrentDate     =   39573
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   405
         TabIndex        =   35
         Top             =   840
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MOTIVO TRASLADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   8640
      TabIndex        =   14
      Top             =   2040
      Width           =   4455
      Begin MSDataListLib.DataCombo DtcMotivoTraslado 
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GUIA DE REMISION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   8640
      TabIndex        =   8
      Top             =   960
      Width           =   4455
      Begin VB.TextBox TxtSerie_Guia 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   11
         Top             =   495
         Width           =   885
      End
      Begin VB.TextBox TxtNumero_Guia 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   10
         Top             =   495
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DtpFecha_Guia 
         Height          =   375
         Left            =   2865
         TabIndex        =   9
         Top             =   435
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   175964161
         CurrentDate     =   39573
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1455
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.TextBox TxtPtoLlegada 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   3
      Top             =   2685
      Width           =   5085
   End
   Begin VB.TextBox TxtPtoPartida 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   100
      TabIndex        =   2
      Top             =   2280
      Width           =   5085
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5280
      Top             =   8640
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
            Picture         =   "FrmGuiaRemision.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmGuiaRemision.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   10380
      TabIndex        =   0
      Top             =   8490
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2715
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
         TabIndex        =   1
         Top             =   30
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   1429
         ButtonWidth     =   1323
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
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Eliminar)"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ORIGEN-DESTINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   6495
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pto.Llegada:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pto.Partida:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATOS DESTINATARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   360
      TabIndex        =   16
      Top             =   3480
      Width           =   6495
      Begin VB.TextBox TxtRucDestinatario 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   20
         Top             =   720
         Width           =   1965
      End
      Begin VB.TextBox TxtRazonSocialDestinatario 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   18
         Top             =   360
         Width           =   5085
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   315
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razon Social:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATOS TRANSPORTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   360
      TabIndex        =   32
      Top             =   6720
      Width           =   5895
      Begin VB.TextBox TxtMTC 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2145
         MaxLength       =   50
         TabIndex        =   39
         Top             =   360
         Width           =   2085
      End
      Begin VB.TextBox TxtPlaca 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2130
         MaxLength       =   60
         TabIndex        =   38
         Top             =   720
         Width           =   2085
      End
      Begin VB.TextBox TxtMarca 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   37
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   45
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado MTC:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   720
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1410
         TabIndex        =   40
         Top             =   1080
         Width           =   525
      End
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   9615
      Left            =   0
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label lblempresa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EL LIRIO DE LOS VALLES SAC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   195
      TabIndex        =   36
      Top             =   960
      Width           =   7305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GUIA DE REMISION "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7440
      TabIndex        =   4
      Top             =   0
      Width           =   4305
   End
End
Attribute VB_Name = "FrmDetalleGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim configura As Boolean

Private Sub Check1_Click()

End Sub









Private Sub Form_Activate()
If Me.TxtNombreEmpresaTransporte.Text <> "" Then
    Me.TxtRuc_Transportes.SetFocus
Else
    Me.TxtCodigoEmpresaTransporte.SetFocus
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.LblEmpresa.Caption = KEY_EMPRESA

 
 
strCadena = "SELECT * FROM DetalleGuia WHERE (sSerieGuia='" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "' AND sNumeroGuia='" & Trim(FrmVentas.TxtNumeroDoc) & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        Call ActualizaComponentes
        Call LLenarGuia
        Exit Sub
    Else
        Call ActualizaComponentes
        Call LlenarGuiaNueva
End If

End Sub
Private Sub ActualizaComponentes()
strCadena = "SELECT cMotivo as Codigo,sMotivo as Descripcion FROM MotivoTransferencia ORDER BY sMotivo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMotivoTraslado)
Me.DtcMotivoTraslado.BoundText = "01"
Set rst = Nothing

strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc_Ref)
  Me.DtcTipoDoc_Ref.BoundText = FrmVentas.DtcComprobanteGuia.BoundText
  Set rst = Nothing







End Sub
Private Sub LLenarGuia()
strCadena = "SELECT * FROM DetalleGuia WHERE (sSerieGuia='" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "'AND " & _
            "sNumeroGuia='" & Trim(FrmVentas.TxtNumeroDoc.Text) & "' AND Alm_cod='" & Trim(FrmVentas.DtcAlmacen.BoundText) & "')"
            
            
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtSerie_Guia.Text = rst!sSerieGuia
    Me.TxtNumero_guia.Text = rst!sNumeroGuia
    Me.DtpFecha_Guia.Value = CVDate(rst!sFechaGuia)
    Me.TxtPtoPartida.Text = rst!sOrigen
    Me.TxtPtoLlegada.Text = rst!sDestino
    Me.DtcMotivoTraslado.Text = rst!sMotivo
    Me.DtcMotivoTraslado.Enabled = False
    Me.TxtRazonSocialDestinatario.Text = rst!sRazonDestinatario
    Me.TxtRucDestinatario.Text = rst!sRucDestinatario
    Me.DtcTipoDoc_Ref.Text = rst!sDoc_codRef
    Me.txtserie_factura.Text = rst!sSerieRef
    Me.txtnumero_factura.Text = rst!sNumeroref
    Me.DtpFecha_Factura.Value = CVDate(rst!sFechaRef)
    Me.TxtCodigoEmpresaTransporte.Text = rst!sCodTransporte
    Me.TxtNombreEmpresaTransporte.Text = rst!sEmpresaTransporte
    Me.TxtDireccionTransporte.Text = rst!sDireccionTransporte
    Me.TxtRuc_Transportes.Text = rst!sRucTransporte
    Me.txtmtc.Text = rst!MTC
    Me.TxtPlaca.Text = rst!Placa
    Me.TxtMarca.Text = rst!marca
    Me.TxtNombreChofer.Text = rst!Chofer
    Me.TxtLicencia.Text = rst!sLicencia
    Me.DtcMotivoTraslado.Enabled = False
    Me.DtcTipoDoc_Ref.Enabled = False
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
    Set rst = Nothing
End If

Set rst = Nothing
End Sub
Private Sub LlenarGuiaNueva()
Me.TxtSerie_Guia.Text = FrmVentas.DtcSerieDoc.BoundText
Me.TxtNumero_guia.Text = FrmVentas.TxtNumeroDoc.Text
Me.DtpFecha_Guia.Value = FrmVentas.DtpActual.Value
Me.TxtPtoPartida.Text = UCase(KEY_DIRECCION)
Me.TxtPtoLlegada.Text = FrmVentas.txtDireccion.Text
Me.TxtRazonSocialDestinatario.Text = FrmVentas.txtCliente.Text
Me.TxtRucDestinatario.Text = FrmVentas.TxtCodCliente.Text
Me.txtserie_factura.Text = FrmVentas.TxtSeri_guia.Text
Me.txtnumero_factura.Text = FrmVentas.TxtNumero_guia.Text
Me.DtpFecha_Factura.Value = FrmVentas.DtpFechaReferencia.Value

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_SAVE
        Call Save
    Case KEY_DELETE
    
        If MsgBox("Esta Seguro de Eliminar esta Guia", vbInformation + vbYesNo, "Mensaje para el Usuario") = vbYes Then
            strCadena = "DELETE FROM DetalleGuia WHERE (sSerieGuia='" & Trim(Me.TxtSerie_Guia.Text) & "' AND sNumeroGuia='" & Trim(Me.TxtNumero_guia.Text) & "' AND Doc_Cod='" & KEY_GUIA & "')"
            Call EjecutaRST(strCadena)
            Unload Me
        End If
    Case KEY_EXIT
        Unload Me
End Select
End Sub
Private Sub Save()
If Trim(Me.txtmtc.Text) <> "" And Trim(Me.txtmtc.Text) <> "TRAMITE" And Trim(Me.TxtPlaca.Text) <> "" And Trim(Me.TxtMarca.Text) <> "" Then
    Call SaveMTC(Trim(Me.txtmtc.Text), Trim(Me.TxtPlaca.Text), Trim(Me.TxtMarca.Text))
End If
If Trim(Me.TxtNombreChofer.Text) <> "" And Trim(Me.TxtNombreChofer.Text) <> "--------------------" And Trim(Me.TxtLicencia.Text) <> "" Then
    Call SaveChofer(Trim(Me.TxtNombreChofer.Text), Trim(Me.TxtLicencia.Text))
End If


 strCadena = "INSERT INTO DetalleGuia(sSerieGuia,sNumeroGuia,doc_cod,Alm_cod,sFechaGuia,sOrigen,sDestino,sMotivo,sRazonDestinatario, " & _
" sRucDestinatario,sDoc_codRef,sSerieRef,sNumeroref,sFechaRef,sCodTransporte,sEmpresaTransporte,sDireccionTransporte," & _
"sRucTransporte,MTC,IntGuia,chofer,placa,marca,slicencia ) VALUES ('" & Trim(Me.TxtSerie_Guia.Text) & "','" & Trim(Me.TxtNumero_guia.Text) & "'," & _
"'" & Trim(FrmVentas.DtcTipoDoc.BoundText) & "','" & Trim(FrmVentas.DtcAlmacen.BoundText) & "','" & CVDate(Me.DtpFecha_Guia.Value) & "','" & Trim(Me.TxtPtoPartida.Text) & "','" & Trim(Me.TxtPtoLlegada.Text) & "'," & _
"'" & Trim(Me.DtcMotivoTraslado.Text) & "','" & Trim(Me.TxtRazonSocialDestinatario.Text) & "','" & Trim(Me.TxtRucDestinatario.Text) & "'," & _
"'" & Trim(Me.DtcTipoDoc_Ref.Text) & "','" & Trim(Me.txtserie_factura.Text) & "','" & Trim(Me.txtnumero_factura.Text) & "'," & _
"'" & CVDate(Me.DtpFecha_Factura.Value) & "','" & Trim(Me.TxtCodigoEmpresaTransporte.Text) & "','" & Trim(Me.TxtNombreEmpresaTransporte.Text) & "'," & _
"'" & Trim(Me.TxtDireccionTransporte.Text) & "','" & Trim(Me.TxtRuc_Transportes.Text) & "'," & _
"'" & Trim(Me.txtmtc.Text) & "','" & GeneraCodigoGuia & "','" & Trim(Me.TxtNombreChofer.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.TxtMarca.Text) & "','" & Trim(Me.TxtLicencia.Text) & "')"




Call EjecutaRST(strCadena)

Set rst = Nothing
Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
End Sub
Private Function GeneraCodigoGuia() As Integer
strCadena = "SELECT intGuia FROM DetalleGuia ORDER BY intGuia Desc"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    GeneraCodigoGuia = rst(0) + 1
Else
    GeneraCodigoGuia = 1
End If
Set rst = Nothing
End Function
Private Sub SaveMTC(ByVal sMTC As String, Placa As String, marca As String)
strCadena = "SELECT MTC FROM UnidadTransporte WHERE MTC='" & sMTC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    Set rst = Nothing
Else
    strCadena = "INSERT INTO UnidadTransporte(MTC,Placa,Marca)VALUES('" & Trim(Me.txtmtc.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.TxtMarca.Text) & "')"
    Call EjecutaRST(strCadena)
    Set rst = Nothing
End If
End Sub
Private Sub SaveChofer(ByVal Nombre As String, sLicencia As String)
strCadena = "SELECT Licencia FROM Chofer WHERE Licencia='" & sLicencia & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    Set rst = Nothing
Else
    strCadena = "INSERT INTO Chofer(Licencia,sChofer)VALUES('" & sLicencia & "','" & Nombre & "')"
    Call EjecutaRST(strCadena)
    Set rst = Nothing
End If
End Sub
Private Sub TxtCodigoEmpresaTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmPersona.Show

End If
End Sub

Private Sub TxtLicencia_Change()
If configura = False Then
    

strCadena = "SELECT * FROM Chofer WHERE Licencia LIKE '%" & Trim(Me.TxtLicencia.Text) & "%' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtNombreChofer.Text = rst("sChofer")
Else
    Me.TxtNombreChofer.Text = ""
End If
Set rst = Nothing
End If
End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    strCadena = "SELECT sChofer FROM Chofer WHERE Licencia='" & Trim(Me.TxtLicencia.Text) & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         Me.TxtNombreChofer.Text = rst(0)
    Else
        Me.TxtNombreChofer.SetFocus
    End If
    Set rst = Nothing
End If
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtLicencia.SetFocus
End If
End Sub

Private Sub TxtMTC_KeyPress(KeyAscii As Integer)
If configura = False Then


If Len(Me.txtmtc.Text) < 1 Then
    Me.TxtPlaca.Text = ""
    Me.TxtMarca.Text = ""
    Exit Sub
End If
strCadena = "SELECT * FROM UnidadTransporte WHERE MTC LIKE '%" & Trim(Me.txtmtc.Text) & "%' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtPlaca.Text = rst("Placa")
    Me.TxtMarca.Text = rst("Marca")
Else
    Me.TxtPlaca.Text = ""
    Me.TxtMarca.Text = ""
End If
Set rst = Nothing
End If
End Sub

Private Sub TxtNombreChofer_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtMarca.SetFocus
End If
End Sub

Private Sub TxtRuc_Transportes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtmtc.SetFocus
End If
End Sub

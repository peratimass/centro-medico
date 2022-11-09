VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmMovimientoCaja 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7275
      MaxLength       =   80
      TabIndex        =   33
      Text            =   "0000000000"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6270
      MaxLength       =   80
      TabIndex        =   32
      Text            =   "0000"
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame7 
      Caption         =   "Comprobante Relacionado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   29
      Top             =   2400
      Width           =   3615
      Begin VB.TextBox TxtComprobanteRelacionado 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   30
         Top             =   360
         Width           =   3285
      End
      Begin MSDataListLib.DataCombo DtcProveedor 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
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
   End
   Begin VB.Frame Frame6 
      Caption         =   "Glosa:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   5175
      Begin VB.TextBox TxtGlosa 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   240
         Width           =   4965
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Pagar con Cheque:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      TabIndex        =   22
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton cmdCargarCheque 
         Caption         =   "Cargar Cheque"
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptChequeNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptChequeSi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Si"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DtcCheque 
         Height          =   315
         Left            =   1680
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Image ImgEscaneo 
         Height          =   375
         Left            =   240
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   5175
      Begin VB.TextBox TxtIdCompra 
         Height          =   375
         Left            =   3240
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtItf 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   19
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox TxtComision 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   18
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITF:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   285
         TabIndex        =   21
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comision:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   285
         TabIndex        =   20
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   5175
      Begin VB.TextBox Txtdepositado 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   16
         Top             =   1080
         Width           =   2325
      End
      Begin VB.TextBox TxtMovimiento 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   15
         Top             =   720
         Width           =   2325
      End
      Begin VB.TextBox txtTC 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   14
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.Depositado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   13
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.Movimiento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cambio :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nº Operacion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
      Begin VB.TextBox txtOperacion 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         MaxLength       =   8
         TabIndex        =   9
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione Cuenta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5175
      Begin MSDataListLib.DataCombo DtcCuentas 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtpOperacion 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   195
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121176065
      CurrentDate     =   40750
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   8760
      Top             =   1800
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
            Picture         =   "FrmMovCaja.frx":0000
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovCaja.frx":0460
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovCaja.frx":04ED
            Key             =   "(Grabar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   6420
      TabIndex        =   0
      Top             =   5175
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2595
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   2475
         _ExtentX        =   4366
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
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   675
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121176065
      CurrentDate     =   40750
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   3240
      TabIndex        =   34
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   330
      Left            =   6240
      TabIndex        =   35
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   975
      Left            =   3120
      Top             =   240
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Valor :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. Operacion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   165
      TabIndex        =   3
      Top             =   240
      Width           =   1125
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1155
      Left            =   120
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "FrmMovimientoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCargarCheque_Click()
On Error GoTo salir
Me.CommonDialog1.Filter = "*.jpg"
Me.CommonDialog1.ShowOpen

Me.ImgEscaneo.Picture = LoadPicture(Me.CommonDialog1.FileName)
img = Me.CommonDialog1.FileName
Exit Sub
salir:
MsgBox "Imagen Incorrecta", vbInformation, "Mensaje para el Usuario"
End Sub

Private Sub DtcCuentas_Change()
strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Val(Me.DtcCuentas.BoundText) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If rst("tipo_cuenta") = "caja" Then
        Me.TxtItf.Text = 0
    Else
        Me.TxtItf.Text = Format(Val(Me.TxtDepositado.Text) * (0.005) / 100, "#,##0.000")
    End If
End If
Set rst = Nothing
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Me.OptChequeNO.Value = True
Dim tItf As Single
Dim tComision As Single
Dim comprobante As String

Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = False
            
   strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
   Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='V' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0096"
  Me.TxtSerie.Text = "0003"
  
  Set rst = Nothing
   strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY numero DESC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtNumeroDoc.Text = rst(0)
    Else
        Me.TxtNumeroDoc.Text = GeneraCodigo(6)
    End If
    Set rst = Nothing
    
    
  
  strCadena = "SELECT mis_cuentas.id_cuenta AS Codigo, (mis_cuentas.descripcion+'-'+ Moneda.descripcion+'-'+mis_cuentas.numero_cuenta) AS Descripcion " & _
  "FROM  mis_cuentas INNER JOIN Moneda ON mis_cuentas.tipo_moneda = Moneda.id_moneda"
   
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentas)
  Me.DtcCuentas.BoundText = 1
  Set rst = Nothing
    
  
  
  Me.TxtTc.Text = KEY_CAMBIO
  Me.TxtIdCompra.Text = Val(FrmListadoFacturasCompra.HfgFacturas.TextMatrix(FrmListadoFacturasCompra.HfgFacturas.Row, 9))
  strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.TxtIdCompra.Text) & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.TxtMovimiento.Text = rst("saldo")
    Me.TxtDepositado.Text = rst("saldo")
    Me.TxtComprobanteRelacionado.Text = FrmListadoFacturasCompra.HfgFacturas.TextMatrix(FrmListadoFacturasCompra.HfgFacturas.Row, 3)
    cPersona = rst("cPersona")
    Set rst = Nothing
    
    strCadena = "SELECT cPersona as Codigo, NombrePersona as Descripcion FROM Persona WHERE cPersona='" & Trim(cPersona) & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
  Set rst = Nothing
  End If
  
  
  tItf = 0
  Me.TxtItf.Text = Format(tItf, "#,##0.000")
End Sub
Private Sub Save()
Dim Saldo As Single
Dim monto_letras As String
Dim Monto As Double
Dim glosaITF As String
  On Error GoTo salir
  If (Me.DtcCuentas.BoundText <> "" Or Val(Me.TxtDepositado.Text) <= 0) Then
    strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.TxtIdCompra.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Saldo = rst("saldo") - Val(Me.TxtDepositado.Text)
        strCadena = "UPDATE DocumentoCompra SET saldo='" & Val(Saldo) & "' WHERE idCompra='" & Val(Me.TxtIdCompra.Text) & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "INSERT INTO mis_cuentas_det(id_cuenta,fecha,fecha_sys,tipo_trans,cPersona,Persona,glosa,monto,montoreal,tc,documento,operacion) " & _
        " VALUES('" & Val(Me.DtcCuentas.BoundText) & "','" & CVDate(Me.DtpValor.Value) & "','" & CVDate(Date) & "','E','" & Trim(Me.DtcProveedor.BoundText) & "','" & Trim(Me.DtcProveedor.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtDepositado.Text) & "','" & Val(Me.TxtDepositado.Text) * -1 & "','" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtComprobanteRelacionado.Text) & "','" & Trim(Me.txtOperacion.Text) & "')"
        CnBd.Execute (strCadena)
         
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
        CnBd.Execute (strCadena)
         
        Monto = Val(Me.TxtDepositado.Text)
        monto_letras = UCase(EnLetras(Monto))
        If Me.OptChequeSi.Value = True Then
        
            strCadena = "SELECT * FROM Cheques WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_cuenta='" & Val(Me.DtcCuentas.BoundText) & "'"
            Call ConfiguraRst(strCadena)
            strCadena = "UPDATE  Cheques SET monto='" & Val(Me.TxtDepositado.Text) & "',cPersona='" & Trim(Me.DtcProveedor.BoundText) & "',Persona='" & Trim(Me.DtcProveedor.Text) & "',fecha='" & CVDate(Me.DtpValor.Value) & "'," & _
            "estado='emitido',detalle='" & Trim(Me.TxtGlosa.Text) & "' WHERE id_cheque='" & Trim(Me.DtcCheque.BoundText) & "' AND id_chequera='" & Val(rst("id_chequera")) & "'             "
            CnBd.Execute (strCadena)
             
            Set rst = Nothing
            strCadena = "INSERT INTO OrdenPago(doc_cod,serie,numero,empresa,direccion,ruc_emp,num_cheque,cpersona,persona,entidad_financiera,fecha,cambio,glosa,monto,monto_letras)VALUES " & _
            " ('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(KEY_EMPRESA) & "','" & Trim(KEY_DIRECCION) & "'," & _
            "'" & Trim(KEY_RUC) & "','" & Trim(Me.DtcCheque.BoundText) & "','" & Trim(Me.DtcProveedor.BoundText) & "','" & Trim(Me.DtcProveedor.Text) & "','" & Trim(Me.DtcCuentas.Text) & "','" & CVDate(Me.DtpValor.Value) & "'," & _
            "'" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtDepositado.Text) & "','" & Trim(monto_letras) & "')"
        Else
            strCadena = "INSERT INTO OrdenPago(doc_cod,serie,numero,empresa,direccion,ruc_emp,num_cheque,cpersona,persona,entidad_financiera,fecha,cambio,glosa,monto,monto_letras)VALUES " & _
            " ('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(KEY_EMPRESA) & "','" & Trim(KEY_DIRECCION) & "'," & _
            "'" & Trim(KEY_RUC) & "','-----','" & Trim(Me.DtcProveedor.BoundText) & "','" & Trim(Me.DtcProveedor.Text) & "','" & Me.DtcCuentas.Text & "','" & CVDate(Me.DtpValor.Value) & "'," & _
            "'" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtGlosa.Text) & "','" & Val(Me.TxtDepositado.Text) & "','" & Trim(monto_letras) & "')"
        End If
          CnBd.Execute (strCadena)
           
          
        If Val(Me.TxtItf.Text) > 0 Then
            glosaITF = "ITF-" + Trim(Me.TxtComprobanteRelacionado.Text)
            strCadena = "INSERT INTO mis_cuentas_det(id_cuenta,fecha,fecha_sys,tipo_trans,cPersona,Persona,glosa,monto,montoreal,tc,documento,operacion) " & _
            " VALUES('" & Val(Me.DtcCuentas.BoundText) & "','" & CVDate(Me.DtpValor.Value) & "','" & CVDate(Date) & "','E','" & Trim(Me.DtcProveedor.BoundText) & "','" & Trim(Me.DtcProveedor.Text) & "','" & Trim(glosaITF) & "','" & Val(Me.TxtItf.Text) & "','" & Val(Me.TxtItf.Text) * -1 & "','" & Val(Me.TxtTc.Text) & "','" & Trim(Me.TxtComprobanteRelacionado.Text) & "','" & Trim(Me.txtOperacion.Text) & "')"
            CnBd.Execute (strCadena)
             
            
        End If
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
        Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
        
        Call FrmListadoFacturasCompra.facturas
    End If
  End If
  
  
  Exit Sub
salir:
  MsgBox "Ocurrio un Error al Grabar", vbInformation, "Mensaje para el Usuario"
  
End Sub

Private Sub OptChequeSi_Click()
If Me.OptChequeSi.Value = True Then
     strCadena = "SELECT id_cheque as Codigo, id_cheque as Descripcion FROM Cheques WHERE id_cuenta='" & Val(Me.DtcCuentas.BoundText) & "' AND estado='libre'"
     Call ConfiguraRst(strCadena)
     Call LlenaDataCombo(Me.DtcCheque)
     Set rst = Nothing
     Me.DtcCheque.Visible = True
     Me.cmdCargarCheque.Visible = True
    Else
     Me.DtcCheque.Visible = False
     Me.cmdCargarCheque.Visible = False
End If

End Sub
Private Sub Imprimir()
strCadena = "SELECT empresa, direccion, ruc_emp, serie, numero, num_cheque, entidad_financiera, cpersona, persona, fecha, cambio, glosa, monto, monto_letras " & _
"FROM  OrdenPago WHERE doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptOrdenPago", , App.Path + "\Reportes\")
End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_PRINT
        Call Imprimir
    Case KEY_EXIT
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub

End Sub

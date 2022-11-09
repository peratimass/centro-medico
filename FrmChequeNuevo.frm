VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmChequeNuevo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtidCuenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   9960
      MaxLength       =   80
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtidChequera 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8280
      MaxLength       =   80
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox lblcostos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3165
      MaxLength       =   80
      TabIndex        =   24
      Top             =   2100
      Width           =   4305
   End
   Begin VB.TextBox TxtidCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6600
      MaxLength       =   80
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Txtcentrocosto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   21
      Top             =   2100
      Width           =   1095
   End
   Begin VB.TextBox TxtTc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10200
      MaxLength       =   80
      TabIndex        =   20
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TxtMontoMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   6600
      MaxLength       =   80
      TabIndex        =   18
      Top             =   2460
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   1815
      Left            =   2040
      TabIndex        =   12
      Top             =   2800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3201
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   8760
      TabIndex        =   15
      Top             =   1080
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      StartOfWeek     =   230686721
      CurrentDate     =   41130
   End
   Begin VB.TextBox TxtMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   14
      Top             =   2460
      Width           =   4455
   End
   Begin VB.CommandButton CmdQuitar 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7560
      Picture         =   "FrmChequeNuevo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   3
      Top             =   1750
      Width           =   5415
   End
   Begin VB.TextBox TxtRazonsocial 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1400
      Width           =   5415
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2055
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1020
      Width           =   1575
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Height          =   315
      Left            =   6600
      MaxLength       =   80
      TabIndex        =   0
      Text            =   "000000"
      Top             =   480
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   11280
      Top             =   3240
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
            Picture         =   "FrmChequeNuevo.frx":058A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":08A6
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":0D06
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":0D93
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":11F3
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":150F
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":196F
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":1C8B
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":20EB
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":254B
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":2E2B
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":3147
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeNuevo.frx":3463
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   8940
      TabIndex        =   4
      Top             =   4320
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
         TabIndex        =   5
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
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.COSTOS :"
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
      Left            =   930
      TabIndex        =   22
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO:"
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
      Left            =   9075
      TabIndex        =   19
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOTIVO :"
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
      Left            =   1110
      TabIndex        =   17
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO A GIRAR :"
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
      TabIndex        =   16
      Top             =   4920
      Width           =   1365
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   870
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
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
      TabIndex        =   9
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   8760
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   855
      Left            =   240
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "FrmChequeNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChkAdelantado_Click()

End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE FROM cheque_detalle WHERE id_item='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
Call llenarGrid(Me.HfDetalle, Val(Me.TxtidCheque.Text))
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.MonthView1.Value = KEY_FECHA
If FrmComprasPagos.Procedencia = nuevo Then
    Call llenar_cheque(Val(FrmComprasPagos.DtcCheque.BoundText))
End If
If FrmCheques.Procedencia = nuevo Then
    Call llenar_cheque(FrmCheques.Hfcheques.TextMatrix(FrmCheques.Hfcheques.Row, 0))
End If
If FrmSolicitudViaticosAtender.Procedencia = nuevo Then
    Call llenar_cheque(Val(FrmSolicitudViaticosAtender.DtcCheque.BoundText))
End If
End Sub
Public Sub llenar_cheque(ByVal id_cheque As Double)
strCadena = "SELECT C.id_estado,C.id_beneficiario,ccostos,fecha_hora,tc,numero,M.id_cuenta,C.id_cheque,CH.id_chequera FROM cheque C,chequera CH,mis_cuentas M WHERE C.id_cheque='" & id_cheque & "' AND C.ruc='" & KEY_RUC & "' AND " & _
" C.id_chequera=CH.id_chequera AND CH.ruc='" & KEY_RUC & "' AND CH.id_cuenta=M.id_cuenta AND M.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst("id_estado") = "02" Then
    Me.TxtNumeroDoc.Text = rst("numero")
    Me.txtruc.Text = rst("id_beneficiario")
    Me.txtrazonsocial.Text = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_beneficiario"))
    Me.txtdireccion.Text = BDBuscarCampo("persona", "direccion", "dni", rst("id_beneficiario"))
    Me.Txtcentrocosto.Text = rst("ccostos")
    Me.lblcostos.Text = BDBuscarCampo("plan_contable_det", "plan_des", "pc_codigo", rst("ccostos"))
    Me.MonthView1.Value = rst("fecha_hora")
    Me.TxtTc.Text = rst("tc")
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
Else
    Me.TxtidCheque.Text = rst("id_cheque")
    Me.TxtidChequera.Text = rst("id_chequera")
    Me.TxtidCuenta.Text = rst("id_cuenta")
    Me.TxtNumeroDoc.Text = rst("numero")
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
     
End If


Me.TxtTc.Text = KEY_CAMBIO
'Me.Caption = UCase(rst("descripcion")) + ":" + Space(2) + UCase(FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 5)) + Space(2) + ":" + UCase(FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 4))
strCadena = "SELECT M.id_cuenta as Codigo,CONCAT(M.descripcion,'-',MO.descripcion,'-',M.numero_cuenta) as Descripcion FROM mis_cuentas M,moneda MO WHERE M.id_moneda=MO.id_moneda  AND M.id_tipo<>'01'AND M.ruc='" & KEY_RUC & "' AND id_cuenta='" & rst("id_cuenta") & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentas)
Call llenarGrid(Me.HfDetalle, id_cheque)



End Sub

Private Sub Save()
Dim Saldo As Double, atendido As String * 2
If Me.txtruc.Text <> "" And Val(Me.txtMonto.Text) > 0 Then
    strCadena = "UPDATE cheque SET monto='" & Val(Me.txtMonto.Text) & "',saldo='" & Val(Me.txtMonto.Text) & "',id_beneficiario='" & Trim(Me.txtruc.Text) & "',fecha_hora='" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "',id_estado='02',dni_save='" & KEY_USUARIO & "',seleccionado='no',ccostos='" & Trim(Me.Txtcentrocosto.Text) & "' WHERE id_cheque='" & Val(Me.TxtidCheque.Text) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    
    strCadena = "UPDATE cheque SET seleccionado='si' WHERE numero='" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "' AND ruc='" & KEY_RUC & "' AND id_chequera='" & Val(Me.TxtidChequera.Text) & "'"
    CnBd.Execute (strCadena)
     
    strCadena = "UPDATE chequera SET cheque_actual='" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "'WHERE id_chequera='" & Val(Me.TxtidChequera.Text) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    'Call FrmCheques.actualizar(Me.HfDetalle, Val(Me.TxtidChequera.text))
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
    If FrmSolicitudViaticosAtender.Procedencia = nuevo Then
        Saldo = Val(FrmSolicitudViaticosAtender.txtsaldo.Text) - Val(Me.txtMonto.Text)
        If Val(FrmSolicitudViaticosAtender.txtsaldo.Text) > Val(Me.txtMonto.Text) Then
            atendido = "no"
        Else
            atendido = "si"
        End If
        strCadena = "UPDATE solicitud_dinero SET saldo='" & Saldo & "',monto_entregado='" & Val(Me.txtMonto.Text) & "',fecha_confirmacion='" & KEY_FECHA & "',hora_confirmacion='" & str(Time) & "',atendido='" & atendido & "' WHERE id_solicitud='" & Val(FrmSolicitudViaticosAtender.Txtid_solicitud.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        FrmSolicitudViaticosAtender.Procedencia = Neutro
        Call FrmSolicitudViaticos.actualizar
    End If
Else
    MsgBox "LLENE TODOS LOS PARAMETROS", vbInformation, KEY_EMPRESA
End If

End Sub



Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub lblcostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
         Procedencia = buscar
         FrmPlanContableCuentas.TxtDescripcion.Text = Trim(Me.lblcostos.Text)
         FrmPlanContableCuentas.Show
         FrmPlanContableCuentas.TxtPlanContable.SetFocus
         Exit Sub
End If
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_EXIT
         If Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True Then
            strCadena = "DELETE FROM cheque_factura WHERE id_cheque='" & Val(Me.TxtidCheque.Text) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
         End If
         Unload Me
         
   
    
  Exit Sub
error:
  
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub



Private Sub Txtcentrocosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
         Procedencia = buscar
         FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.Txtcentrocosto.Text)
         FrmPlanContableCuentas.Show
         FrmPlanContableCuentas.TxtPlanContable.SetFocus
         Exit Sub

End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(Me.txtMonto.Text) > 0 Then
    strCadena = "SELECT sum(montoreal) FROM mis_cuentas_det WHERE id_cuenta='" & Val(FrmChequeras.HfgDetalle.TextMatrix(FrmChequeras.HfgDetalle.Row, 1)) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = True Then
        If MsgBox("NO CUENTA CON SALDO EN LA CUENTA" + Chr(13) + "DESEA CONTINUAR", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
            Me.txtMonto.Text = Format(Val(Me.txtMonto.Text), "###0.00")
            Call Resalta(Me.Txtcentrocosto)
            Exit Sub
        Else
            Call Resalta(Me.txtMonto)
            Exit Sub
        End If
    Else
        If Val(Me.txtMonto.Text) > rst(0) Then
            If MsgBox("EL MONTO A GIRAR EXEDE AL DISPONIBLE" + Chr(13) + "DESEA CONTINUAR", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                Call Resalta(Me.Txtcentrocosto)
                Exit Sub
            Else
                Call Resalta(Me.txtMonto)
                Exit Sub
            End If
        Else
            Me.txtMonto.Text = Format(Val(Me.txtMonto.Text), "###0.00")
            Call Resalta(Me.Txtcentrocosto)
        End If
    End If
End If
End Sub

Private Sub TxtMontoMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.TxtMontoMotivo.Text) > 0 And Len(Me.txtMotivo.Text) > 0 Then
        strCadena = "INSERT INTO cheque_detalle(id_cheque,detalle,monto,ruc)VALUES('" & Val(Me.TxtidCheque.Text) & "','" & UCase(Me.txtMotivo.Text) & "','" & Val(Me.TxtMontoMotivo.Text) & "','" & KEY_RUC & "')"
        Call CnBd.Execute(strCadena)
        Call llenarGrid(Me.HfDetalle, Val(Me.TxtidCheque.Text))
        Me.txtMotivo.Text = ""
        Me.TxtMontoMotivo.Text = ""
        Call Resalta(Me.txtMotivo)
        
    End If
End If
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal id_cheque As Double)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM cheque_detalle WHERE id_cheque='" & id_cheque & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.txtMonto.Text = 0
    Grilla.Rows = 1
    Me.CmdQuitar.Visible = False
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
           Grilla.ColWidth(2) = 3400
           Grilla.ColWidth(3) = 1200
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
       Me.txtMonto.Text = Format(tTotal, "###0.00")
      For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
      Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(Me.txtMotivo.Text) > 0 Then
    Call Resalta(Me.TxtMontoMotivo)
    Exit Sub
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtruc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtrazonsocial.Text = UCase(rst("nombre_completo"))
        Me.txtdireccion.Text = UCase(rst("direccion"))
        Call Resalta(Me.Txtcentrocosto)
        Exit Sub
    Else
        If Me.txtruc.Text <> "" Then
            If MsgBox("Cliente No Registrado Desea registrarlo Ahora ??", vbInformation + vbYesNo, "Diguite el ruc del Usuario .") = vbYes Then
                Procedencia = 1
                FrmDetallePersona.txtruc.Text = Trim(Me.txtruc.Text)
                FrmDetallePersona.Show
                Exit Sub
            End If
         Else
           Procedencia = buscar
           FrmPersona.Show
           Exit Sub
                End If
        End If
        
    End If
 

End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleComprobante 
   BorderStyle     =   0  'None
   Caption         =   "Detalle Comprobantes:"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodComprobante 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      MaxLength       =   80
      TabIndex        =   36
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DFDFE0&
      Height          =   735
      Left            =   2040
      TabIndex        =   33
      Top             =   1680
      Width           =   3855
      Begin VB.CheckBox ChkTienda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Utilizado en Tienda."
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   35
         Top             =   240
         Width           =   1770
      End
      Begin VB.CheckBox ChkAfectaStock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Afecta a Stock."
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DFDFE0&
      Caption         =   "Documento Sunat"
      Height          =   855
      Left            =   6120
      TabIndex        =   30
      Top             =   1560
      Width           =   1695
      Begin VB.OptionButton OptNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptSI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox TxtPor2Impuesto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7320
      MaxLength       =   80
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox TxtPor1Impuesto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7320
      MaxLength       =   80
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox TxtAbvreviatura 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      MaxLength       =   80
      TabIndex        =   23
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox TxtDesCentroCostos 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3120
      MaxLength       =   80
      TabIndex        =   22
      Top             =   6600
      Width           =   6135
   End
   Begin VB.TextBox TxtDesAsientoNormal 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3120
      MaxLength       =   80
      TabIndex        =   21
      Top             =   6120
      Width           =   6135
   End
   Begin VB.TextBox TxtGlosaLibMayor 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3120
      MaxLength       =   80
      TabIndex        =   20
      Top             =   5640
      Width           =   6135
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      MaxLength       =   80
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4680
      Top             =   7200
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
            Picture         =   "FrmDetalleComprobante.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleComprobante.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   8100
      TabIndex        =   1
      Top             =   7215
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
         TabIndex        =   2
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
   Begin MSDataListLib.DataCombo DtcCtaContableTotal 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcCta1Impuesto 
      Height          =   315
      Left            =   2760
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcCta2Impuesto 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcCtaNeto 
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DTCAfectaNeto 
      Height          =   315
      Left            =   7320
      TabIndex        =   16
      Top             =   4440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   10455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glosas para Asientos Contables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   600
      TabIndex        =   29
      Top             =   5040
      Width           =   3945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Contables e Importes de Prevision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   600
      TabIndex        =   28
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de Identificacion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   675
      TabIndex        =   27
      Top             =   120
      Width           =   2865
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje 2º Impuesto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5400
      TabIndex        =   26
      Top             =   4080
      Width           =   1725
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Centro de Costos:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   780
      TabIndex        =   19
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label LblAsirntoNormal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Asiento Normal:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   780
      TabIndex        =   18
      Top             =   6120
      Width           =   2025
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glosa Libro Mayor:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   780
      TabIndex        =   17
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   600
      Top             =   5400
      Width           =   9375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "El % Afecta al NETO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5355
      TabIndex        =   15
      Top             =   4560
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje 1º Impuesto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5385
      TabIndex        =   14
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Ctable del Neto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Ctable del 2º Impuesto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Ctable del 1º Impuesto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   705
      TabIndex        =   11
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cta Contable del Total:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   705
      TabIndex        =   7
      Top             =   3120
      Width           =   1635
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2100
      Left            =   600
      Top             =   2880
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abreviatura:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   795
      TabIndex        =   6
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label LblLaboratorio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   945
      TabIndex        =   4
      Top             =   645
      Width           =   555
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1980
      Left            =   600
      Top             =   480
      Width           =   9375
   End
End
Attribute VB_Name = "FrmDetalleComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodComprobante As String
Dim DocTienda As String * 1


Private Sub ChkTienda_Click()
If Me.ChkTienda.Value = 1 Then
    DocTienda = "V"
Else
    DocTienda = "F"
End If
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Top = 500
End Sub
Private Sub Form_Load()
'---llenar los data Combos-----


  Select Case FrmComprobantes.Procedencia
    Case nuevo
         
    Case Modificar
          Call LLENA
  End Select
End Sub
Private Sub LLENA()
    strCadena = "SELECT * FROM comprobantes WHERE id_doc = '" & Trim(FrmComprobantes.HfgComprobantes.TextMatrix(FrmComprobantes.HfgComprobantes.Row, 0)) & "'"
  Call EjecutaRST(strCadena)
  
  Me.txtCodComprobante.Text = RstEjecuta!id_doc
  'Format(Trim(Str(Val(Right(RstEjecuta!doc_cod, 4)))), "0000")
  Me.txtdescripcion.Text = RstEjecuta!doc_des
  Me.TxtAbvreviatura.Text = RstEjecuta!doc_abrev
  Me.TxtPor1Impuesto.Text = RstEjecuta!doc_impto1
  Me.TxtPor2Impuesto.Text = RstEjecuta!doc_impt2
  
  If (RstEjecuta!doc_Tienda) = "V" Then
    Me.ChkTienda.Value = 1
  End If

If (RstEjecuta!sunat) = "si" Then
    Me.OptSi.Value = True
Else
    Me.OptNo.Value = True
End If
  
    
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub
Private Sub Save()
Dim sunat As String

  If Me.TxtAbvreviatura.Text = "" Or Me.txtdescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
      
  Me.txtCodComprobante.Text = formato_item(Me.txtCodComprobante.Text, 4)
  If Me.OptSi.Value = True Then
    sunat = "si"
    Else
    sunat = "no"
  End If
     
    Select Case FrmComprobantes.Procedencia
     Case nuevo
                    
    strCadena = "SELECT * FROM comprobantes WHERE id_doc='" & Trim(Me.txtCodComprobante.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
    
    strCadena = "INSERT INTO comprobantes(id_doc,doc_des,doc_abrev,sunat) VALUES " & _
    "('" & Trim(Me.txtCodComprobante.Text) & "','" & Trim(Me.txtdescripcion.Text) & "','" & Trim(Me.TxtAbvreviatura.Text) & "','" & sunat & "')"
    CnBd.Execute (strCadena)
     
    Call FrmComprobantes.actualizar
    Unload Me
    Else
        MsgBox "Codigo Comprobante Ya Registrado"
        Set rst = Nothing
        Exit Sub
    End If
            
            
      Case Modificar
            
                strCadena = "UPDATE comprobantes SET doc_des='" & Me.txtdescripcion.Text & "'," & _
                "doc_abrev='" & Me.TxtAbvreviatura.Text & "',doc_tienda=" & _
                " '" & DocTienda & "',sunat='" & sunat & "' WHERE id_doc= '" & Trim(Me.txtCodComprobante.Text) & "'"
                CnBd.Execute (strCadena)
                
               Call FrmComprobantes.actualizar
                
                
        Unload Me
    End Select
  End If

End Sub

Private Sub txtCodComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCodComprobante.Text = formato_item(Me.txtCodComprobante.Text, 4)
    Me.txtdescripcion.SetFocus
End If
End Sub

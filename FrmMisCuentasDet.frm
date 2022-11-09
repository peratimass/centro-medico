VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmMisCuentasDet 
   BorderStyle     =   0  'None
   Caption         =   "Detalle Cuentas"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_detraccion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CUENTA DETRACCION"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   21
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox TxtNumCuenta 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox txtCuentaContable 
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
      Left            =   2160
      MaxLength       =   80
      TabIndex        =   9
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   4455
      Begin VB.OptionButton OptCaja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CAJA"
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "BANCO"
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
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton OptGastos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "GASTOS"
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
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
      Begin VB.OptionButton OptDolares 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DOLARES"
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
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton OptSoles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SOLES"
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtBuscar 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CheckBox chkpasarella 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "FORMA PAGO:"
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
      Height          =   290
      Left            =   480
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6360
      Top             =   2520
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
            Picture         =   "FrmMisCuentasDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMisCuentasDet.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5100
      TabIndex        =   11
      Top             =   5160
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
         TabIndex        =   12
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
   Begin MSDataListLib.DataCombo DtcEntidadBancaria 
      Height          =   330
      Left            =   2160
      TabIndex        =   13
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   ""
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
   Begin MSDataListLib.DataCombo DtcTargeta 
      Height          =   330
      Left            =   2160
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   ""
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CUENTA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   750
      TabIndex        =   20
      Top             =   2100
      Width           =   1125
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTA CONTABLE :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   630
      TabIndex        =   19
      Top             =   300
      Width           =   1245
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTIDAD BANCARIA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   18
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO DE CUENTA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   150
      TabIndex        =   17
      Top             =   3480
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1050
      TabIndex        =   16
      Top             =   2880
      Width           =   825
   End
   Begin VB.Label lblPlan 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   600
      Width           =   5535
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6180
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "FrmMisCuentasDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub chkpasarella_Click()
If Me.chkpasarella.Value = 1 Then
  Me.DtcTargeta.Visible = True
Else
 Me.DtcTargeta.Visible = False
End If
End Sub

Private Sub DtcEntidadBancaria_Change()
If Me.DtcEntidadBancaria.Text <> "CAJA" Then
    Me.OptBanco.Value = True
Else
    Me.OptCaja.Value = True
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM targeta WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTargeta)


strCadena = "SELECT codigo as Codigo, descripcion as Descripcion FROM entidadfinanciera   ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEntidadBancaria)
  



If FrmMiscuentas.Procedencia = modificar Then
    Call llenar
End If
  
  
End Sub
Private Sub llenar()
Dim codigo_entidad As String
Dim cuenta_contable As String
    strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Val(FrmMiscuentas.HfgDetalle.TextMatrix(FrmMiscuentas.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cuenta_contable = rst("cuenta_ctble")
        codigo_entidad = rst("id_entidad")
        
        If rst("id_moneda") = "00001" Then
            Me.OptSoles.Value = True
        Else
            Me.OptDolares.Value = True
        End If
        
        If rst("id_tipo") = "01" Then
            Me.OptCaja.Value = True
        End If
        If rst("id_tipo") = "02" Then
            Me.OptBanco.Value = True
        End If
        If rst("id_tipo") = "03" Then
            Me.OptGastos.Value = True
        End If
        
        If rst("detraccion") = "si" Then
            Me.chk_detraccion.Value = 1
        Else
            Me.chk_detraccion.Value = 0
        End If
        
       Me.TxtNumCuenta.Text = rst("numero_cuenta")
       
       Call load_tarjeta(rst("cuenta_ctble"))
       
       
    End If
    Set rst = Nothing
    Me.DtcEntidadBancaria.BoundText = codigo_entidad
    Me.txtCuentaContable.Text = cuenta_contable
    
End Sub
Private Sub load_tarjeta(ByVal in_cuenta As String)
strCadena = "SELECT * FROM targeta WHERE numero_cuenta='" & Trim(in_cuenta) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.chkpasarella.Value = 1
   Me.DtcTargeta.BoundText = rst("id")
   Me.DtcTargeta.Visible = True
Else
   Me.DtcTargeta.Visible = False
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
Private Sub Save()
  If Me.DtcEntidadBancaria.BoundText = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
   If (Me.OptCaja.Value = True) Then
       tipo_cuenta = "01"
    End If
    If Me.OptBanco.Value = True Then
        tipo_cuenta = "02"
    End If
    If Me.OptGastos.Value = True Then
        tipo_cuenta = "03"
    End If
    If tipo_cuenta = "" Then
        MsgBox "Elija un Tipo de Cuenta a Crear", vbInformation, "Mensaje para el Usuario"
        Exit Sub
    End If
    If Me.OptDolares.Value = True Then
        tipo_moneda = "00002"
    End If
    If Me.OptSoles.Value = True Then
        tipo_moneda = "00001"
    End If
    
    If Me.chkpasarella.Value = 1 Then
       in_tarjeta = Me.DtcTargeta.BoundText
    Else
       in_tarjeta = 0
    End If
    
    If Me.chk_detraccion.Value = 1 Then
       in_detraccion = "si"
    Else
       in_detraccion = "no"
    End If
    
    
    If tipo_moneda = "" Then
        MsgBox "Elija una moneda para la Cuenta ", vbInformation, "Mensaje para el Usuario"
        Exit Sub
    End If
    
    
    Call put_save_tarjeta(Me.DtcTargeta.BoundText)
      
      
      
      Select Case FrmMiscuentas.Procedencia
      Case Nuevos
        
        strCadena = "INSERT INTO mis_cuentas(id_entidad,descripcion,numero_cuenta,id_moneda,id_tipo,cuenta_ctble,detraccion,ruc)VALUES ('" & Trim(Me.DtcEntidadBancaria.BoundText) & "','" & Trim(Me.DtcEntidadBancaria.Text) & "','" & Trim(Me.TxtNumCuenta.Text) & "','" & Trim(tipo_moneda) & "','" & Trim(tipo_cuenta) & "','" & Trim(Me.txtCuentaContable.Text) & "','" & in_detraccion & "','" & KEY_RUC & "')"
        Call Execute_Sql(strCadena)
        
        Call FrmMiscuentas.actualizar
        Unload Me
        Exit Sub
        
      Case modificar
        strCadena = "UPDATE mis_cuentas SET detraccion='" & in_detraccion & "', id_entidad='" & Trim(Me.DtcEntidadBancaria.BoundText) & "',numero_cuenta='" & Trim(Me.TxtNumCuenta.Text) & "',id_moneda='" & Trim(tipo_moneda) & "',id_tipo='" & Trim(tipo_cuenta) & "',cuenta_ctble='" & Trim(Me.txtCuentaContable.Text) & "' WHERE id_cuenta = '" & Val(FrmMiscuentas.HfgDetalle.TextMatrix(FrmMiscuentas.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        Call Execute_Sql(strCadena)
        Unload Me
        FrmMiscuentas.Procedencia = Neutro
        Call FrmMiscuentas.actualizar
        Exit Sub
    End Select
    FrmMiscuentas.Procedencia = Neutro
  End If
End Sub
Private Sub put_save_tarjeta(ByVal in_tarjeta As String)

If Me.chkpasarella.Value = 1 Then
  
   
   strCadena = "UPDATE targeta SET numero_cuenta='" & Trim(Me.txtCuentaContable.Text) & "' WHERE id='" & in_tarjeta & "' and ruc='" & KEY_RUC & "'"
   Call Execute_Sql(strCadena)
Else
   strCadena = "UPDATE targeta SET numero_cuenta='-' WHERE id='" & in_tarjeta & "' and ruc='" & KEY_RUC & "'"
   Call Execute_Sql(strCadena)
End If



End Sub


Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If
End Sub

Private Sub txtBuscar_Change()
  
  strCadena = "SELECT codigo as Codigo, descripcion as Descripcion FROM entidadfinanciera WHERE descripcion LIKE '%" & Trim(Me.txtBuscar.Text) & "%'   ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcEntidadBancaria)
  
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcEntidadBancaria.SetFocus
End If
End Sub

Private Sub txtCuentaContable_Change()
If Trim(Me.txtCuentaContable.Text) <> "" Then
    strCadena = "SELECT * FROM con_cuentacontable WHERE NroCuenta='" & Trim(Me.txtCuentaContable.Text) & "' AND activo=1 AND IdEmpresaSis='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.lblPlan.Caption = rst("Descripcion")
    Else
        Me.lblPlan.Caption = ""
    End If
    Set rst = Nothing
End If
End Sub

Public Sub get_cuenta(ByVal in_cuenta As String)
strCadena = "SELECT * FROM con_cuentacontable WHERE NroCuenta='" & Trim(in_cuenta) & "' AND activo=1 AND IdEmpresaSis='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.lblPlan.Caption = rst("Descripcion")
    Else
        Me.lblPlan.Caption = ""
    End If
    Set rst = Nothing
End Sub


Private Sub txtCuentaContable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCuentaContable.Text)
    Exit Sub
End If
End Sub


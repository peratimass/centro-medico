VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmChequeraDet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle Chequera"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
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
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox txtchqueactual 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtchequeInicial 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtnumcheques 
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
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6120
      Top             =   2880
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
            Picture         =   "FrmChequeraDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmChequeraDet.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   5760
      TabIndex        =   8
      Top             =   1680
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1995
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1376
         ButtonWidth     =   1296
         ButtonHeight    =   1376
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   420
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHEQUE ACTUAL:"
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
      Left            =   435
      TabIndex        =   5
      Top             =   1840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHEQUE INICIAL:"
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
      Left            =   465
      TabIndex        =   4
      Top             =   1380
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO CHEQUES:"
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
      Left            =   285
      TabIndex        =   2
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTA BANCARIA:"
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
      Left            =   255
      TabIndex        =   1
      Top             =   420
      Width           =   1515
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2715
      Left            =   120
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "FrmChequeraDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

CenterForm Me
Me.Top = 500
strCadena = "SELECT M.id_cuenta as Codigo,CONCAT(M.descripcion,'-',MO.descripcion,'-',M.numero_cuenta) as Descripcion FROM mis_cuentas M,moneda MO WHERE M.id_moneda=MO.id_moneda  AND M.id_tipo<>'01'AND M.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentas)
If FrmMiscuentas.Procedencia = Modificar Then
    'Call llena
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
Dim id_chequera As Double, flag As String * 2
  If Me.DtcCuentas.BoundText = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
       
    Select Case FrmChequeras.Procedencia
      Case nuevo
        strCadena = "P_insert_chequera('" & Me.DtcCuentas.BoundText & "','" & formato_item(Me.txtchequeInicial.Text, 6) & " ','" & formato_item(Me.txtchqueactual.Text, 6) & "','" & formato_item(Val(Me.txtchequeInicial.Text) + Val(Me.txtnumcheques.Text), 6) & "','" & Val(Me.txtnumcheques.Text) & "','" & Trim(Me.TxtObservacion.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
       
        id_chequera = LastRegistroRUC("chequera", "id_chequera")
         For i = Val(Me.txtchequeInicial.Text) To (Val(Me.txtnumcheques.Text) + Val(Me.txtchequeInicial.Text))
                
                    If i = Val(Me.txtchqueactual.Text) Then
                        flag = "si"
                    Else
                        flag = "no"
                    End If
                    strCadena = "INSERT INTO cheque(numero,id_chequera,seleccionado,ruc)VALUES('" & formato_item(i, 6) & "','" & id_chequera & "','" & flag & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
        Next i
        
        Call FrmChequeras.actualizar
        Unload Me
      Case Modificar
        
       ' StrCadena = "UPDATE mis_cuentas SET cod_banco='" & Trim(Me.DtcEntidadBancaria.BoundText) & "',numero_cuenta='" & Trim(Me.TxtNumCuenta.Text) & "',tipo_moneda='" & Trim(tipo_moneda) & "',tipo_cuenta='" & Trim(tipo_cuenta) & "' WHERE id_cuenta = '" & Val(FrmMiscuentas.HfgDetalle.TextMatrix(FrmMiscuentas.HfgDetalle.Row, 5)) & "'"
        'CnBd.Execute (strCadena)
        'Call FrmMiscuentas.actualizar
        'Unload Me
                
    End Select
    FrmMiscuentas.Procedencia = Neutro
  End If
End Sub
Private Sub TxtBanco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If
End Sub



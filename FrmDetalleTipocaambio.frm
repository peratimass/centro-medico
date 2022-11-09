VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetalleTipocaambio 
   BorderStyle     =   0  'None
   Caption         =   "Tipo Cambio"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdSbs 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "IMPORTAR SBS"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleTipocaambio.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker Dtp_fecha 
      Height          =   440
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   230096897
      CurrentDate     =   43507
   End
   Begin VB.TextBox TxtId_codigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtLocal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtCompra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtventa 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1560
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
            Picture         =   "FrmDetalleTipocaambio.frx":001C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":0338
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":0798
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":0BF8
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":0F14
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":1374
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":1690
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":1AF0
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":1F50
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":2830
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":2B4C
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipocaambio.frx":2E68
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   2760
      TabIndex        =   1
      Top             =   2760
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1995
      _CBHeight       =   870
      _Version        =   "6.7.9782"
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
         Width           =   1875
         _ExtentX        =   3307
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA :"
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
      Left            =   1350
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR LOCAL [SBS] :"
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
      Left            =   390
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR COMPRA [SBS] :"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA [SBS] :"
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
      Left            =   330
      TabIndex        =   3
      Top             =   900
      Width           =   1635
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3705
      Left            =   0
      Top             =   0
      Width           =   6690
   End
End
Attribute VB_Name = "FrmDetalleTipocaambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim StrCodTabla As String
Dim StrCodUnidad As String

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub CmdActualizar_Click()

End Sub

Private Sub cmdSbs_Click()

 Call get_cambio_sbs(Me.Dtp_fecha.Value)
 
 
 
 strCadena = "SELECT * FROM tipo_cambio WHERE  id_creador='" & KEY_RUC & "' and fecha='" & Format(Me.Dtp_fecha.Value, "YYYY-mm-dd") & "' LIMIT 1"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Call Me.LLENA(rst("id_tipocambio"))
 End If
 
 FrmTipocambio.actualizar


End Sub

Private Sub Form_Activate()
CenterForm Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
  'If FrmTipocambio.Procedencia = modificar Or FrmTipocambio.Procedencia = nuevo Then
  '  Call LLENA(Me.TxtId_codigo.Text)
 ' End If
  
End Sub

Public Sub LLENA(ByVal in_cambio As String)

  strCadena = "SELECT * FROM tipo_cambio WHERE  id_tipocambio='" & Val(in_cambio) & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.Dtp_fecha.Value = Format(rst("fecha"), "dd-mm-YYYY")
    Me.txtCompra.Text = Format(rst("valor_compra"), "#,##0.0000")
    Me.TxtVenta.Text = Format(rst("valor_venta"), "#,##0.0000")
    Me.txtLocal.Text = Format(rst("valor_local"), "#,##0.0000")
    
  End If
End Sub

Private Sub Save()
Dim Valor As Single
Dim fecha As Date
Dim fecha1 As String

  If Val(Me.txtCompra.Text) <= 0 Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    
    
    strCadena = "SELECT * FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "' and fecha='" & Format(Me.Dtp_fecha.Value, "YYYY-mm-dd") & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         strCadena = "UPDATE tipo_cambio SET valor_compra='" & Val(Me.txtCompra.Text) & "',valor_venta='" & Val(Me.TxtVenta.Text) & "',valor_local='" & Val(Me.txtLocal.Text) & "' WHERE   id_tipocambio='" & Val(Me.TxtId_codigo.Text) & "'"
    Else
        strCadena = "INSERT INTO tipo_cambio(descripcion,fecha,valor,valor_venta,valor_compra,valor_venta1,valor_local,id_creador)VALUES" & _
        "('Compra Dolar','" & Format(Me.Dtp_fecha.Value, "YYYY-mm-dd") & "','" & Val(txtCompra.Text) & "','" & Val(Me.TxtVenta.Text) & "','" & Val(Me.txtCompra.Text) & "','" & Val(Me.txtLocal.Text) & "','" & KEY_RUC & "')"
    End If
    CnBd.Execute (strCadena)
    
    fecha1 = Format(FrmTipocambio.HfgCambio.TextMatrix(FrmTipocambio.HfgCambio.Row, 2), "yyyy-mm-dd")
    
    
    
    
    FrmTipocambio.actualizar
    Unload Me
End If
End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
      Unload Me
      Call enabled_form(FrmTipocambio)
      Exit Sub
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub


Private Sub TxtResumen_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub




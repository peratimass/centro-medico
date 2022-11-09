VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteBarras 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impresion Codigos Barra"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12480
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SstKardex 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Producto"
      TabPicture(0)   =   "FrmReporteRegCompras.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBarras"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigoProducto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDescripcion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtPrecio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCopias"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdImpresionContinua"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CmdReporte"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CommandButton CmdReporte 
         Caption         =   "IMPRESION VISUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdImpresionContinua 
         Caption         =   "IMPRESION CONTINUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtCopias 
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
         Left            =   5400
         TabIndex        =   10
         Top             =   1150
         Width           =   1575
      End
      Begin VB.TextBox TxtPrecio 
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
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1150
         Width           =   2940
      End
      Begin VB.TextBox txtCodigoProducto 
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
         Left            =   2040
         TabIndex        =   5
         Top             =   750
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "NUMERO COPIAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblBarras 
         BackStyle       =   0  'Transparent
         Caption         =   "12345"
         BeginProperty Font 
            Name            =   "3 of 9 Barcode"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   2040
         TabIndex        =   8
         Top             =   1920
         Width           =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO BARRA :"
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
         Left            =   600
         TabIndex        =   3
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO VENTA :"
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
         Left            =   645
         TabIndex        =   2
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO INTERNO :"
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
         Left            =   420
         TabIndex        =   1
         Top             =   840
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   240
         Top             =   600
         Width           =   7215
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6720
      Top             =   2655
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
            Picture         =   "FrmReporteRegCompras.frx":001C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":0470
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":0790
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":0BE4
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":1038
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":1358
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":17AC
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":1908
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":1D5C
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":2078
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":2954
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":2C74
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteRegCompras.frx":2F94
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteBarras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()

End Sub

Private Sub cmdImpresionContinua_Click()
Call impresion_barras(Trim(Me.txtCodigoProducto.text), Val(Me.txtCopias.text))
End Sub

Private Sub CmdReporte_Click()
strCadena = "SELECT P.nombre_prod,P.precio_venta,B.cod_barra FROM producto P,producto_barras B WHERE P.id_producto=B.id_producto AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.txtCodigoProducto.text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Ans = ShowMultiReport(rst, "rptBarras", , App.Path + "\Reportes\")
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
End Sub

Private Sub Text1_Change()

End Sub

Public Sub presionar(ByVal id_producto As String)
Dim cproducto As String
Me.txtCodigoProducto.text = formato_item(Trim(id_producto), 5)
 If (Len(id_producto) = 0) Or Val(id_producto) = 0 Then
        
        Call Resalta(Me.txtCodigoProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
    strCadena = "SELECT * FROM producto_barras B,producto P  WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.id_producto='" & Trim(Me.txtCodigoProducto.text) & "'"

    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtCodigoProducto.text = rst("id_producto")
       Me.txtDescripcion.text = UCase(rst("nombre_prod"))
       Me.TxtPrecio.text = Format(rst("precio_venta"), "#,##0.00")
       Me.lblBarras.Caption = rst("cod_barra")
    Else
       Me.txtDescripcion.text = ""
       Me.TxtPrecio.text = ""
       Me.lblBarras.Caption = ""
       Call Resalta(Me.txtCodigoProducto)
    End If
End Sub
Private Sub TxtCodigoProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call presionar(Trim(Me.txtCodigoProducto.text))
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteSalidaDinero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frm Salida de Dinero"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   4
      Top             =   2460
      Width           =   1575
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2100
      Width           =   3975
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1740
      Width           =   3975
   End
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1740
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   870
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   154861569
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   870
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   154861569
      CurrentDate     =   37091
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2865
      Left            =   5205
      TabIndex        =   7
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   5054
      BandCount       =   1
      ForeColor       =   8388608
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   2865
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2430
         Left            =   30
         TabIndex        =   8
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   4286
         ButtonWidth     =   1482
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Imprimir"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSalidaDinero.frx":0000
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSalidaDinero.frx":008D
            Key             =   "(Salir)"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   2085
      TabIndex        =   9
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin VB.CheckBox ChkAlmacen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Almacen"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   270
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   240
      Top             =   75
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Persona:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   255
      TabIndex        =   13
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label LblIdentificacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   285
      TabIndex        =   12
      Top             =   2460
      Width           =   555
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "al"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2445
      TabIndex        =   10
      Top             =   960
      Width           =   225
   End
End
Attribute VB_Name = "FrmReporteSalidaDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.DtpDesde.Value = Date
Me.DtpHasta.Value = Date
strCadena = "SELECT Alm_cod as Codigo,Alm_des as Descripcion FROM Almacen ORDER BY Alm_des"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.Enabled = True
Me.ChkAlmacen.Value = 1
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

  Select Case Button.key
    Case KEY_PRINT
            Call Imprimir(Trim(Me.DtcAlmacen.BoundText), Trim(Me.TxtCodCliente.Text))
    Case KEY_EXIT
        Unload Me
  Exit Sub
End Select
End Sub
Private Sub Imprimir(ByVal CodAlm As String, ByVal Cliente As String)
Dim Ans As Boolean
Dim CodComp As String
CodComp = "0097"
strCadena = "SELECT DocumentoVenta.cPersona, DocumentoVenta.Persona," & _
    "DocumentoVenta.dEmisionVenta,Comprobantes.doc_abrev , (DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta),DocumentoVenta.nTotalVenta  " & _
    "FROM DocumentoVenta INNER JOIN Persona ON DocumentoVenta.cPersona = Persona.cPersona INNER JOIN " & _
    "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Cliente & "' AND DocumentoVenta.Alm_Cod='" & CodAlm & "' AND " & _
    "DocumentoVenta.dEmisionVenta >='" & CVDate(Me.DtpDesde.Value) & "' AND DocumentoVenta.dEmisionVenta <='" & CVDate(Me.DtpHasta.Value) & "' AND Comprobantes.doc_cod='" & Trim(CodComp) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        
       ' Set Me.HfgExcel.Recordset = Rst
        Ans = ShowMultiReport(rst, "RptSalDinero", , App.Path + "\Reportes\")
    Else
        MsgBox "Usuario no cuenta con Salida de Comprobantes", vbInformation
    End If

End Sub


Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtCodCliente.Text = FormatosCeros(Me.TxtCodCliente.Text, 5)
    strCadena = "SELECT cPersona,NombrePersona,sRazonSocial,sDireccionCliente1,Per_Ruc FROM " & _
    " Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Trim(rst(5)) = "N" Then
            Me.TxtCliente.Text = rst(1)
        Else
            Me.TxtCliente.Text = rst(2)
        End If
        
        Me.TxtDireccion.Text = rst(3)
        Me.txtruc.Text = rst(4)
        
    End If
End If
If (KeyAscii = 66 Or KeyAscii = 98) Then
    Procedencia = Selecionar
    FrmPersona.Show
End If
End Sub



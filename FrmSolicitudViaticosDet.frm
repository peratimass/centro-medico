VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmSolicitudViaticosDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   18690
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtId_solicitud 
      Height          =   285
      Left            =   8040
      TabIndex        =   20
      Top             =   360
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   300
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.TextBox TxtCcostos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   2040
      Top             =   4200
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
            Picture         =   "FrmSolicitudViaticosDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSolicitudViaticosDet.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   14940
      TabIndex        =   2
      Top             =   7200
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2835
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   810
      Width1          =   855
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   810
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Solicitar"
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
   Begin MSComCtl2.DTPicker DtpFechaSolicitud 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   122683393
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   122683393
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   315
      Left            =   3480
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   122683393
      CurrentDate     =   37091
   End
   Begin VB.Label lblAdvertencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   9360
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL A SOLICITAR :"
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
      Height          =   210
      Left            =   195
      TabIndex        =   18
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RESUMEN :"
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
      Height          =   210
      Left            =   975
      TabIndex        =   17
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nº SOLICITUD:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4200
      TabIndex        =   15
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION :"
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
      Height          =   210
      Left            =   600
      TabIndex        =   14
      Top             =   5760
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE :"
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
      Height          =   210
      Left            =   1050
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO DE GASTOS :"
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
      Height          =   210
      Left            =   75
      TabIndex        =   12
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label lblccostos 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3480
      TabIndex        =   10
      Top             =   1320
      Width           =   6705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.COSTOS :"
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
      Height          =   210
      Left            =   900
      TabIndex        =   9
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label lblhora 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3480
      TabIndex        =   8
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA SOLICITUD :"
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
      Height          =   210
      Left            =   330
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblSolicitud 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOLICITUD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   5520
      TabIndex        =   6
      Top             =   300
      Width           =   2250
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   13095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "FrmSolicitudViaticosDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idRecibo As Double
Dim StrCodTabla As String
Dim strCodLinea As String
Public Procedencia As EnumProcede












Private Sub CmdQuitar_Click()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 And Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True Then
    strCadena = "DELETE FROM solicitud_dinero_detalle WHERE  ruc='" & KEY_RUC & "' AND id_detalle='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "'"
    CnBd.Execute (strCadena)
    
    Call llenar_detalle(Me.HfDetalle, Val(Me.TxtId_solicitud.Text))
Else
   strCadena = "DELETE FROM solicitud_dinero_temporal WHERE id='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
    
   Call llenarGrid(Me.HfDetalle)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1000
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
 

 
  
  Select Case FrmSolicitudViaticos.Procedencia
     Case nuevo
        Call llenarGrid(Me.HfDetalle)
        Me.lblhora.Caption = Time
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
        strCadena = "SELECT * FROM solicitud_dinero WHERE dni='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND finalizado='no' AND anulado='no'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 1 Then
            Me.lblAdvertencia.Visible = True
            Me.lblAdvertencia.Caption = "USTED YA TIENE 2 SOLICITUDES PENDIENTE DE DECLARAR" + Chr(13) + Chr(13) + "**** LA GERENCIA ****"
            Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
            Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
            Exit Sub
        End If
        
    Case Modificar
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Me.lblSolicitud.Caption = FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 1)
        Call LLENA(FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 0))
     
  End Select
End Sub
 Private Sub Imprimir()
  Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    'Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
   ' strCadena = " SELECT * FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Transporte = MiTransporte.id_Transporte INNER JOIN " & _
    "Persona ON OrdenCompra.cPersona = Persona.cPersona " & _
    " WHERE serie='" & Me.TxtSerie.text & "' AND numero='" & Me.TxtNumero.text & "' AND doc_cod='0110'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Set rst = Nothing
        MsgBox "Imposible de Imprimir, Error en la Impresora", vbInformation, "Mensaje para el Usuario"
        Exit Sub
    End If
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(80); "ORD COMP:"; Space(1); Mid(Me.TxtSerie.text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumero.text
    Printer.Print ""
     Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(40); Me.TxtGlosa.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(20); rst("marca")
      Printer.Print ""
    strCadena = "SELECT     MiTransporteTipo.Descripcion, MiTransporte.marca, MiTransporte.placa FROM  MiTransporte INNER JOIN  MiTransporteTipo ON MiTransporte.id_tipoTransporte = MiTransporteTipo.idTipo WHERE MiTransporte.id_Transporte='" & rst("id_cisterna") & "' AND Ruc='" & KEY_RUC & "'"
    Call ConfiguraTemporal(strCadena)
    Printer.Print Tab(20); rst("placa") + Space(20) + rstTemporal("Descripcion") + Space(1) + rstTemporal("marca") + "-" + rstTemporal("placa")
    Set rstTemporal = Nothing
    strCadena = "SELECT * FROM Persona WHERE cPersona='" & rst("cConductor") & "'"
    Call ConfiguraTemporal(strCadena)
    Printer.Print ""
    Printer.Print Tab(30); rstTemporal("licencia")
    Printer.Print ""
    Printer.Print Tab(25); rstTemporal("NombrePersona")
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(45); str(Day(rst("fecha"))) + Space(5) + str(Month(rst("fecha"))) + Space(10) + str(Year(rst("fecha")))
    Printer.EndDoc
    Exit Sub



 End Sub


Private Sub Save()
Dim id_solicitud As Double, Numero As String
  If Val(Me.txtMonto.Text) <= 0 Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmSolicitudViaticos.Procedencia
      Case nuevo
       
       strCadena = "SELECT * FROM solicitud_dinero WHERE dni='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND atendido='si' AND finalizado='no'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            MsgBox "USUARIO CUENTA CON " + str(rst.RecordCount) + Space(1) + "SOLICITUDES PENDIENTE", vbInformation, "Mensaje para el Usuario"
       End If
        strCadena = "SELECT * FROM solicitud_dinero WHERE ruc='" & KEY_RUC & "' ORDER BY id_solicitud DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            Numero = formato_item(Val(rst("numero")) + 1, 6)
        Else
            Numero = formato_item(1, 6)
        End If
        strCadena = "INSERT INTO solicitud_dinero (numero,resumen,monto_solicitado,saldo,fecha_solicitud,hora_solicitud,ccostos,fecha_inicio,fecha_fin,observacion,dni,ruc) VALUES " & _
        " ('" & Numero & "','" & UCase(Trim(Me.txtResumen.Text)) & "','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & KEY_FECHA & "','" & str(Time) & "','" & Trim(Me.TxtCcostos.Text) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'," & _
        "'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtObservacion.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        id_solicitud = LastRegistro("solicitud_dinero", "id_solicitud")
        strCadena = "SELECT * FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO solicitud_dinero_detalle(id_solicitud,descripcion,monto,ruc)VALUES('" & id_solicitud & "','" & rst("detalle") & "','" & rst("monto") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                rst.MoveNext
           Next i
           strCadena = "DELETE FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
            
        End If
        
        Call FrmSolicitudViaticos.actualizar
       Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
       Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
       End Select
   End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
     Case KEY_PRINT
        Call Imprimir
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub


Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub






Private Sub TxtCodProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     If Trim(TxtcodProveedor.Text) = "" Then
        Procedencia = buscar
        FrmPersona.Show
        Exit Sub
    Else
        
       ' If Len(Me.TxtCodProveedor.Text) < 8 Then
           ' strCadena = "SELECT * FROM Persona WHERE cPersona='" & Trim(Me.TxtCodProveedor.Text) & "'"
        'Else
         '   strCadena = "SELECT * FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtCodProveedor.Text) & "'"
        'End If
        'Call ConfiguraRst(strCadena)
        'If rst.RecordCount > 0 Then
         '  Me.TxtCodProveedor.Text = rst("cPersona")
           'Me.TxtProveedor.Text = rst("NombrePersona")
          ' Me.DtcTipoTransporte.SetFocus
          ' Set rst = Nothing
        'Else
         '   Procedencia = buscar
          '  FrmPersona.Show
           ' Exit Sub
        End If
    End If
'End If
End Sub



Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtCodProveedor)
End If
End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ' Call Resalta(Me.TxtAutorizacion)
End If
End Sub

Public Sub LlenarTelefonos(ByVal Grilla As MSHFlexGrid, ByVal cPersona As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT idTelefono,telefono FROM  PersonaTelefono WHERE cPersona='" & cPersona & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 700
            Grilla.ColWidth(1) = 2500
            
        Next
        cabecera = "CODIGO" & vbTab & "TELEFONO"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("idTelefono") & vbTab & rst("telefono")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
     
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If

End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_PRINT
             strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE codigo='" & idRecibo & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
    
  Exit Sub
    Case KEY_EXIT
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub


Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoMotivo)
End If
End Sub

Private Sub TxtMonto_Change()
If Val(Me.txtMonto.Text) > 0 Then
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
Else
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
End If
End Sub


Private Sub TxtCcostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
         Procedencia = buscar
         FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.TxtCcostos.Text)
         FrmPlanContableCuentas.Show
         FrmPlanContableCuentas.TxtPlanContable.SetFocus
         Exit Sub
End If
End Sub

Private Sub TxtMontoMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If Val(Me.TxtId_solicitud.Text) > 0 Then
        strCadena = "INSERT INTO solicitud_dinero_detalle(id_solicitud,descripcion,monto,ruc)VALUES('" & Val(Me.TxtId_solicitud.Text) & "','" & UCase(Me.txtdetalle.Text) & "','" & Val(Me.TxtMontoMotivo.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Call llenar_detalle(Me.HfDetalle, Val(Me.TxtId_solicitud.Text))
        Me.txtdetalle.Text = ""
        Me.TxtMontoMotivo.Text = 0
        Call FrmSolicitudViaticos.actualizar
        Call Resalta(Me.txtdetalle)
        Exit Sub
    End If
    
   
End If
End If
End Sub
Private Sub llenar_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_solicitud As Double)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM solicitud_dinero_detalle WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
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
             Fila = rst("id_detalle") & vbTab & formato_item(i, 2) & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "#,##0.00")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL A SOLICITAR  ***********" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       Me.txtMonto.Text = Format(tTotal, "###0.00")
      For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
      Next k
      strCadena = "UPDATE solicitud_dinero SET monto_solicitado='" & Val(Me.txtMonto.Text) & "' WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
       
      Call FrmSolicitudViaticos.actualizar
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
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
             Fila = rst("id") & vbTab & formato_item(i, 2) & vbTab & rst("detalle") & vbTab & Format(rst("monto"), "#,##0.00")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL A SOLICITAR  ***********" & vbTab & Format(tTotal, "###0.00")
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



Private Sub TxtResumen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdetalle)
End If
End Sub

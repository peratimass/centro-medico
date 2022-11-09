VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmIngresoDinero 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frm Ingreso de Dinero"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   495
      MaxLength       =   80
      TabIndex        =   10
      Top             =   1380
      Width           =   735
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      MaxLength       =   80
      TabIndex        =   9
      Top             =   1380
      Width           =   4695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1320
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2460
      Width           =   4695
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      MaxLength       =   80
      TabIndex        =   7
      Top             =   1740
      Width           =   4695
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      MaxLength       =   80
      TabIndex        =   6
      Top             =   2100
      Width           =   1695
   End
   Begin VB.TextBox TxtMontoIngresar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3390
      MaxLength       =   80
      TabIndex        =   0
      Text            =   "0000"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4395
      MaxLength       =   80
      TabIndex        =   4
      Text            =   "0000000000"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox ChkSalidaDinero 
      Caption         =   "Listado de Ingresos de Dinero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   3135
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Informe"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPDesde 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   135593985
      CurrentDate     =   39788
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1560
      Top             =   3000
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
            Picture         =   "FrmIngresoDinero.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmIngresoDinero.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   4020
      TabIndex        =   11
      Top             =   3120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   2430
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   4286
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
               Caption         =   "Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Imprimir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpActual 
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Top             =   2085
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   135593985
      CurrentDate     =   37091
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   600
      TabIndex        =   14
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   12648447
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   360
      Left            =   3360
      TabIndex        =   15
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   12648447
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   3225
      Left            =   6360
      TabIndex        =   16
      Top             =   240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   5689
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3225
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   4050
         Left            =   30
         TabIndex        =   17
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   7144
         ButtonWidth     =   1773
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  &Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2415
      Left            =   195
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   7215
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135593985
         CurrentDate     =   39788
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         FillStyle       =   5  'Downward Diagonal
         Height          =   615
         Left            =   120
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   840
         Width           =   135
      End
      Begin VB.Label LblMonto 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   5325
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   1560
      End
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
      Left            =   510
      TabIndex        =   29
      Top             =   1080
      Width           =   1425
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
      Left            =   360
      TabIndex        =   28
      Top             =   1740
      Width           =   915
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
      Left            =   645
      TabIndex        =   27
      Top             =   2100
      Width           =   435
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs:"
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
      Left            =   585
      TabIndex        =   26
      Top             =   2580
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   25
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   3495
      TabIndex        =   24
      Top             =   2160
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   480
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmIngresoDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Check1_Click()

End Sub

Private Sub ChkSalidaDinero_Click()
If Me.ChkSalidaDinero.Value = 1 Then
    Me.HfdDetalle.Visible = True
    strCadena = "SELECT DocumentoVenta.dEmisionVenta as FECHA, (Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as NUMERO, " & _
    "DocumentoVenta.nTotalVenta as MONTO , Seguridad.Usuario as ENTREGO FROM DocumentoVenta INNER JOIN " & _
    "Seguridad ON DocumentoVenta.id_usuario = Seguridad.IdUsuario INNER JOIN " & _
    "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
    "WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND DocumentoVenta.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'ORDER BY 1 ASC"
    Call LLenaGrilla
Else
    Me.HfdDetalle.Clear
    Me.HfdDetalle.Visible = False
    Me.cmdimprimir.Visible = False
End If
End Sub
Private Sub LLenaGrilla()
Dim Monto As Double
Dim i As Integer
Monto = 0
     Call ConfiguraRst(strCadena)
    Me.HfdDetalle.Clear
     If rst.RecordCount > 0 Then
        Me.cmdimprimir.Visible = True
        Me.LblMonto.Visible = True

        Set Me.HfdDetalle.Recordset = rst
        HfdDetalle.ColWidth(0) = 1100
        HfdDetalle.ColWidth(1) = 2500
        HfdDetalle.ColWidth(2) = 1300
        HfdDetalle.ColWidth(3) = 1600
        Call DarFormatoFecha(HfdDetalle, 0)
        Call DarFormato(HfdDetalle, 2)
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Monto = Monto + rst(2)
             rst.MoveNext
        Next i
        Me.LblMonto.Caption = "S/." + Space(2) + Format(str(Monto), "#,##0.00")
     End If
    'Set Rst = Nothing
End Sub

Private Sub CmdFiltrar_Click()
strCadena = "SELECT DocumentoVenta.dEmisionVenta as FECHA, (Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as NUMERO, " & _
    "DocumentoVenta.nTotalVenta as MONTO , Seguridad.USUARIO as ENTREGO FROM DocumentoVenta INNER JOIN " & _
    "Seguridad ON DocumentoVenta.IdUsuario = Seguridad.IdUsuario INNER JOIN " & _
    "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
    "WHERE (DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND DocumentoVenta.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND DocumentoVenta.dEmisionVenta >='" & CVDate(DtpDesde.Value) & "' AND DocumentoVenta.dEmisionVenta <='" & CVDate(Me.DtpHasta.Value) & "' )ORDER BY 1 ASC"
    Call LLenaGrilla
End Sub

Private Sub cmdimprimir_Click()
    Dim TotalP As String
    Dim Total As Double
    Dim fecha As String
    Dim Documento As String
    Dim usuario As String
    Dim totalletras As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "COD      :" + Space(1); Me.TxtCodCliente.Text
    Printer.Print Tab(4); "CLIENTE  :" + Space(1); Mid(Me.TxtCliente.Text + Space(80), 1, 40) & Space(1) & CVDate(Me.DtpActual.Value)
    Printer.Print ""
    Printer.Print Tab(4); "DIRECCIÓN:" + Space(1); Mid(Me.txtdireccion.Text + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(4); "REPORTE AL:" + Space(1) + str(Date) + Space(2) + "A LAS:" + Space(1) + str(Time)
    Printer.Print "----------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    'StrCadena = "SELECT DocumentoVenta.dEmisionVenta as FECHA, (Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as NUMERO, " & _
    '"DocumentoVenta.nTotalVenta*-1 as MONTO , Seguridad.USUARIO as ENTREGO FROM DocumentoVenta INNER JOIN " & _
    '"Seguridad ON DocumentoVenta.IdUsuario = Seguridad.IdUsuario INNER JOIN " & _
    '"Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
    '"WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND DocumentoVenta.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'ORDER BY 1 ASC"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    Total = 0
            For j = 0 To rst.RecordCount - 1
                fecha = CVDate(rst(0))
                Documento = Mid(rst(1) + Space(30), 1, 30)
                Total = Total + rst(2)
                TotalP = Mid(Format(str(rst(2)), "#,##0.00") + Space(4), 1, 8)
                usuario = Mid(rst(3) + Space(30), 1, 20)
                Printer.Print Tab(1); fecha & Space(3) & Documento & Space(3) & TotalP & Space(4) & usuario
                Printer.CurrentY = Printer.CurrentY + 0.4
                rst.MoveNext
            Next j
            inc = 0.5
            Printer.Print Tab(40); "==================="
            Printer.Print Tab(40); Format(str(Total), "#,##0.00")
            Printer.Print Tab(40); "==================="
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    rst.MoveFirst
    Total = Format(str(Total), "#,##0.00")
    totalletras = UCase(EnLetras(Format(str(Total), "#,##0.00")))
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
    Printer.Print Tab(40); Mid(Total & Space(20), 1, 11)
    Printer.Print "----------------------------------------------------------------------"
    Printer.EndDoc
    
    Exit Sub
    
    
     

    
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcTipoDoc.BoundText) <> "0001" And Trim(Me.DtcTipoDoc.BoundText) <> "0003" Then
        Call Resalta(Me.TxtSerie)
    Else
        MsgBox "Srta:" + Space(1) + KEY_VENDEDOR + Space(1) + "Solo Salida de Dinero", vbInformation
        Me.DtcTipoDoc.SetFocus
    End If
    
    
End If
End Sub

Public Sub nuevo()
    If Me.DtcAlmacen.Enabled = True Then
    
    strCadena = "SELECT cDocumentoVenta FROM DocumentoVenta WHERE (sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY intDocumentoVenta DESC"
    Call ConfiguraRst(strCadena)
    Me.TxtNumeroDoc.Text = GeneraCodigo(10)
    Me.TxtCodCliente.Text = "00000"
    Me.TxtCliente.Text = ""
    Me.txtdireccion.Text = ""
    Me.txtobservacion.Text = ""
    Me.TxtMontoIngresar.Text = "0.00"
    Me.ChkSalidaDinero.Value = 0
    Me.LblMonto.Visible = False
    Me.TxtRuc.Text = ""
    Me.DtcAlmacen.SetFocus
    Set rst = Nothing
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
Else
    MsgBox "Active el Almacen Correspondiente", vbInformation, KEY_EMPRESA
End If
End Sub

Private Sub Form_Activate()
'Call Resalta(Me.TxtSerie)

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
doc_Tienda = "V"
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
   Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='" & doc_Tienda & "' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0016"
  Me.TxtSerie.Text = "0003"
  Set rst = Nothing
  Me.DtpActual.Value = Date
  Me.DtpDesde.Value = CVDate(Date)
  Me.DtpHasta.Value = CVDate(Date)
End Sub
Private Sub LlenarDatosCliente(ByVal Numero As String, ByVal Documento As String, ByVal serie As String, ByVal Almacen As String)
Dim CodPersona As String
strCadena = "SELECT cPersona,Persona,dEmisionVenta,nTotalVenta,Observacion FROM DocumentoVenta WHERE (cDocumentoVenta ='" & Numero & "' AND doc_cod='" & Documento & "' AND sSerie='" & serie & "' AND Alm_Cod='" & Almacen & "')"
Call ConfiguraRst(strCadena)
    CodPersona = Trim(rst(0))
    Me.TxtCodCliente.Text = CodPersona
    Me.TxtCliente.Text = rst(1)
    Me.DtpActual.Value = CVDate(rst(2))
    Me.TxtMontoIngresar.Text = Format(rst(3), "#,##0.00")
    On Error GoTo SALIR
    Me.txtobservacion.Text = rst(4)
    'Me.DtpFechaReferencia.Value = CVDate(Rst(2))
    'Me.DtcFormaPago.BoundText = Rst(3)
SALIR:
Set rst = Nothing
strCadena = "SELECT sDireccionCliente1,Per_Ruc FROM Persona WHERE (cPersona ='" & CodPersona & "' )"
Call ConfiguraRst(strCadena)

        Me.txtdireccion.Text = rst(0)
        Me.TxtRuc.Text = rst(1)
        

Set rst = Nothing
Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True

End Sub
Private Sub Save()
Dim i As Integer
Dim MontoIngresar As Double
Dim anul As String * 1
Dim Contado As String
Contado = "0001"
anul = "F"

MontoIngresar = Me.TxtMontoIngresar.Text
'01----------------guardar en Documento venta---------------------
nuevo_documento = GeneraCodigoVenta(10)

strCadena = "INSERT INTO DocumentoVenta(id_documentoventa,cDocumentoVenta,doc_cod,Alm_cod,sSerie,cPersona,Persona,Observacion,idFormaPago," & _
            "dEmisionVenta,nTotalVenta,FechaProceso,intDocumentoVenta,Anulado,id_usuario)" & _
            "VALUES ('" & nuevo_documento & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "'," & _
            "'" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "','" & Trim(Me.txtobservacion.Text) & "','" & Trim(Contado) & "'," & _
            "'" & Me.DtpActual.Value & "','" & MontoIngresar & "','" & CVDate(Date) & "','" & Val(Me.TxtNumeroDoc.Text) & "','" & anul & "','" & KEY_USUARIO & "')"
            Call EjecutaRST(strCadena)
            Set rst = Nothing
            Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
            Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
            Exit Sub
End Sub
Private Function GeneraCodigoVenta(ByVal longitud As Integer) As String
Dim x As Integer
Dim rst_v As New ADODB.Recordset

strCadena = "SELECT * FROM DocumentoVenta WHERE sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'  ORDER BY id_documentoventa DESC"
        
rst_v.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic

        
Dim Formato As String
  Formato = ""
  For x = 1 To longitud
    Formato = Formato + "0"
  Next x
   
  If (rst_v.BOF And rst_v.EOF) Then
    StrNumero = Format(str(Val(Formato) + 1), Formato)
  Else
    StrNumero = Format(Trim(str(Val(Right(rst_v(0), longitud + 1)) + 1)), Formato)
  End If
  Set rst = Nothing
  GeneraCodigoVenta = Gencodigo + StrNumero
  Gencodigo = ""
Set rst_v = Nothing
End Function

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
      Call nuevo
    Case KEY_DELETE
       If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_PRINT
      Call Imprimir(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
    
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub

Private Sub Imprimir(ByVal TipoDoc As String, ByVal CodAlm As String, ByVal serie As String, ByVal Numero As String)
Dim i As Integer, j As Integer
Dim totalletras As String

    Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 10
       
If Me.DtcTipoDoc.BoundText = KEY_INGDINER Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); str(CVDate(Me.DtpActual.Value)) + Space(3) + str(Time)
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); "CLIENTE  :" + Space(1) + Mid(Me.TxtCliente.Text + Space(80), 1, 65)
    Printer.Print Tab(5); "DIRECCION:" + Space(1) + Mid(Me.txtdireccion.Text + Space(80), 1, 65)
    Printer.Print Tab(5); "RUC      :" + Space(1) + Mid(Me.TxtRuc.Text + Space(50), 1, 30) & "INGDINER"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); "Monto Efectivo:" & "=============" & Space(20) & Me.TxtMontoIngresar.Text
    Printer.CurrentY = Printer.CurrentY + 10
    totalletras = UCase(EnLetras(Me.TxtMontoIngresar.Text))
    Set rst = Nothing
    '---- fin totales
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 60)
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(60); Me.TxtMontoIngresar.Text
    Printer.EndDoc
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Exit Sub
End If
End Sub
Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtCodCliente.Text = FormatosCeros(Me.TxtCodCliente.Text, 5)
    strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
    " Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Trim(rst(6)) = "J" Then
            Me.TxtCliente.Text = rst(2)
        Else
            Me.TxtCliente.Text = rst(1)
        End If
        
        Me.txtdireccion.Text = rst(3)
        Me.TxtRuc.Text = rst(4)
        Me.txtobservacion.Text = rst(5)
        Call Resalta(Me.txtobservacion)
    End If
End If
If (Me.TxtCodCliente.Text = "00000" And KeyAscii = 13) Then
    Procedencia = Selecionar
    FrmPersona.Show
End If
End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub TxtMontoIngresar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMontoIngresar.Text = Format(Me.TxtMontoIngresar.Text, "#,##0.00")
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 10)
    strCadena = "SELECT cDocumentoVenta,sSerie FROM DocumentoVenta WHERE (cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    
    If rst.RecordCount < 1 Then
       
        If CVDate(Me.DtpActual.Value) <> Date Then
            If MsgBox("la Fecha no Coincide con la Fecha del Documento...Desea Continuar", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                Me.TxtCodCliente.Text = "00000"
                Call Resalta(Me.TxtCodCliente)
            Else
                Me.DtpActual.SetFocus
            End If
        Else
        Me.TxtCodCliente.Text = "00000"
        Call Resalta(Me.TxtCodCliente)
          End If
    Else
        MsgBox "Documento ya Existe", vbInformation, KEY_EMPRESA
       Call LlenarDatosCliente(Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.DtcAlmacen.BoundText))
       Call Resalta(Me.TxtNumeroDoc)
        
        
    End If
End If
Set rst = Nothing
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoIngresar)
End If
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = FormatosCeros(Me.TxtSerie.Text, 4)
    strCadena = "SELECT Alm_cod,doc_cod,serie FROM Det_alm_com WHERE (Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'" & _
        " AND  doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Me.TxtNumeroDoc.Text = "" Then
            
            Me.TxtNumeroDoc.SetFocus
        Else
            Set rst = Nothing
            strCadena = "SELECT cDocumentoVenta FROM DocumentoVenta WHERE (doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "') ORDER BY intDocumentoVenta DESC"
            Call ConfiguraRst(strCadena)
            Me.TxtNumeroDoc.Text = GeneraCodigo(10)
            Call Resalta(Me.TxtNumeroDoc)
        End If
    Else
        MsgBox "Serie no Asiganda a a dicho Almacen", vbInformation, KEY_EMPRESA
    End If
End If

End Sub


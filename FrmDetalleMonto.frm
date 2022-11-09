VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleMonto 
   BorderStyle     =   0  'None
   Caption         =   "Montos Cliente:"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkSalidaDinero 
      Caption         =   "Listado de Salidas de Dinero:"
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
      TabIndex        =   14
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Informe"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxtMontoNuevo 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TxtMontoIngresar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox TxtMontoActual 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1680
      Top             =   2040
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
            Picture         =   "FrmDetalleMonto.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleMonto.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   4620
      TabIndex        =   8
      Top             =   1320
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
         TabIndex        =   9
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
   Begin MSComCtl2.DTPicker DTPDesde 
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3120
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdAdelantos 
      Height          =   1575
      Left            =   315
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   2778
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
      Height          =   6135
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   7215
      Begin VB.Frame Frame3 
         Caption         =   "Salidas de Dinero:"
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   6975
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSalidas 
            Height          =   1695
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   2990
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
         Begin VB.Label LblSalidas 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   23
            Top             =   2040
            Width           =   2025
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Adelantos:"
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   6975
         Begin VB.Label LblAdelantos 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4680
            TabIndex        =   22
            Top             =   1845
            Width           =   2055
         End
      End
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
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
      Begin VB.Label Label6 
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
         TabIndex        =   19
         Top             =   840
         Width           =   135
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Cliente :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Nuevo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Ingresar"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Actual:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   2175
      Left            =   120
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "FrmDetalleMonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total_Adelantos As Double
Dim Total_Salidas As Double

Private Sub ChkSalidaDinero_Click()
    
If Me.ChkSalidaDinero.Value = 1 Then
 
    strCadena = "SELECT IdAdelanto as NUM , Fecha,Cliente, MontoAdelantado " & _
    "FROM DetalleAdelantos " & _
    "WHERE DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "' ORDER BY 2 ASC"
    Call LLenaGrilla
    Call LlenarGrillaSalidas
    Me.TxtMontoActual.Text = Format(Total_Adelantos - Total_Salidas, "#,##0.00")
Else
    Me.HfdAdelantos.Clear
    Me.HfdAdelantos.Visible = False
    Me.cmdimprimir.Visible = False
End If
End Sub

Private Sub CmdFiltrar_Click()
Dim Anulado As String
Anulado = "F"
strCadena = "SELECT IdAdelanto as NUM , Fecha,Cliente, MontoAdelantado " & _
    "FROM DetalleAdelantos " & _
    "WHERE DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "' AND DetalleAdelantos.Fecha>='" & CVDate(Me.DtpDesde.Value) & "' AND DetalleAdelantos.Fecha<='" & CVDate(Me.DtpDesde.Value) & "' ORDER BY 2 ASC"
    Call LLenaGrilla
    
    strCadena = "SELECT DocumentoVenta.dEmisionVenta as Fecha,(Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta)Numero, " & _
            "DocumentoVenta.cPersona , DocumentoVenta.nTotalVenta FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND " & _
            "DocumentoVenta.idFormaPago='" & KEY_ADELANTO & "' AND DocumentoVenta.dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND DocumentoVenta.dEmisionVenta <= '" & CVDate(Me.DtpHasta.Value) & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            Me.HfgSalidas.Clear
        End If
         If rst.RecordCount > 0 Then
        Me.cmdimprimir.Visible = True
        Me.HfgSalidas.Visible = True
        Me.HfgSalidas.Clear
        Set Me.HfgSalidas.Recordset = rst
        HfgSalidas.ColWidth(0) = 1050
        HfgSalidas.ColWidth(1) = 3000
        HfgSalidas.ColWidth(2) = 1000
        HfgSalidas.ColWidth(3) = 1500
        Call DarFormatoFecha(HfgSalidas, 0)
        Call DarFormato(HfgSalidas, 3)
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Monto = Monto + rst(3)
             rst.MoveNext
        Next i
        Me.LblSalidas.Caption = "S/." + Space(2) + Format(str(Monto), "#,##0.00")
     End If
    Set rst = Nothing
    
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Top = 200
Call LLENA
End Sub
Private Sub LLenaGrilla()
Dim Monto As Double
Dim i As Integer
Monto = 0
     Call ConfiguraRst(strCadena)
        Total_Adelantos = 0
     If rst.RecordCount > 0 Then
        Me.cmdimprimir.Visible = True
        Me.HfdAdelantos.Visible = True
        Me.HfdAdelantos.Clear
        Set Me.HfdAdelantos.Recordset = rst
        HfdAdelantos.ColWidth(0) = 500
        HfdAdelantos.ColWidth(1) = 1300
        HfdAdelantos.ColWidth(2) = 3000
        HfdAdelantos.ColWidth(3) = 1600
        Call DarFormatoFecha(HfdAdelantos, 1)
        Call DarFormato(HfdAdelantos, 3)
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Monto = Monto + rst(3)
             rst.MoveNext
        Next i
        Me.LblAdelantos.Caption = "S/." + Space(2) + Format(str(Monto), "#,##0.00")
        Total_Adelantos = Monto
     End If
    Set rst = Nothing
End Sub
Private Sub LlenarGrillaSalidas()
Dim Anulado As String
Anulado = "F"
strCadena = "SELECT DocumentoVenta.dEmisionVenta as Fecha,(Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta)Numero, " & _
            "DocumentoVenta.cPersona as Persona , DocumentoVenta.nTotalVenta AS Total FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND " & _
            "DocumentoVenta.idFormaPago='" & KEY_ADELANTO & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "'"
        Call ConfiguraRst(strCadena)
         Total_Salidas = 0
         If rst.RecordCount > 0 Then
        Me.cmdimprimir.Visible = True
        Me.HfgSalidas.Visible = True
        Me.HfgSalidas.Clear
        Set Me.HfgSalidas.Recordset = rst
        HfgSalidas.ColWidth(0) = 1050
        HfgSalidas.ColWidth(1) = 2900
        HfgSalidas.ColWidth(2) = 900
        HfgSalidas.ColWidth(3) = 1500
        Call DarFormatoFecha(HfgSalidas, 0)
        Call DarFormato(HfgSalidas, 3)
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Monto = Monto + rst(3)
             rst.MoveNext
        Next i
        Me.LblSalidas.Caption = "S/." + Space(2) + Format(str(Monto), "#,##0.00")
        Total_Salidas = Monto
     End If
            
End Sub

Private Sub Form_Load()
Me.DtpDesde.Value = CVDate(Date)
Me.DtpHasta.Value = CVDate(Date)
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

Private Sub LLENA()
FrmPersona.HfdPersona.col = 0
strCadena = "SELECT cPersona,NombrePersona,MontoAdelantado FROM Persona WHERE cPersona='" & Trim(FrmPersona.HfdPersona.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtCodCliente.Text = rst(0)
    Me.TxtCliente.Text = rst(1)
    Me.TxtMontoActual.Text = "Pulse Chek" 'Format(Rst(4), "#,##0.00")
    Call Resalta(Me.TxtMontoIngresar)
End If
Set rst = Nothing
End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub
Private Sub Save()
    Dim CodAdelanto As Double
    Dim Adelantado As Double
    Dim NuevoMonto As Double
    Dim MontoActual As Double
    Dim CodDetalle As Double
    MontoActual = CDbl(Me.TxtMontoNuevo.Text)
    NuevoMonto = CDbl(Me.TxtMontoIngresar.Text)
        strCadena = "UPDATE Persona SET MontoAdelantado='" & MontoActual & "',Fecha='" & CVDate(Date) & "' " & _
        " WHERE cPersona = '" & Trim(Me.TxtCodCliente.Text) & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
    strCadena = "SELECT IdAdelanto FROM DetalleAdelantos ORDER BY IdAdelanto DESC "
    Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        CodAdelanto = 1
    Else
        CodAdelanto = rst(0) + 1
    End If
    
    
    Set rst = Nothing
    strCadena = "INSERT INTO DetalleAdelantos(IdAdelanto,IdCliente,Cliente,Fecha,MontoAdelantado)" & _
    "VALUES ('" & CodAdelanto & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Me.TxtCliente.Text & "'," & _
    "'" & CVDate(Date) & "','" & NuevoMonto & "')"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    'actualizar
    strCadena = "SELECT IdAdelanto as NUM , Fecha,Cliente, MontoAdelantado " & _
    "FROM DetalleAdelantos " & _
    "WHERE DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "' ORDER BY 2 ASC"
    Call LLenaGrilla
    Call LlenarGrillaSalidas
    Me.TxtMontoActual.Text = Format(Total_Adelantos - Total_Salidas, "#,##0.00")
  
  
End Sub

Private Sub TxtMontoIngresar_KeyPress(KeyAscii As Integer)
Dim MontoActual As Double
Dim MontoIngresar As Double
Dim NuevoMonto As Double

If KeyAscii = 13 Then
    MontoActual = Me.TxtMontoActual.Text
    MontoIngresar = Me.TxtMontoIngresar.Text
    NuevoMonto = MontoActual + MontoIngresar
    Me.TxtMontoIngresar.Text = Format(MontoIngresar, "#,##0.00")
     
    Me.TxtMontoNuevo.Text = Format(NuevoMonto, "#,##0.00")
End If
End Sub

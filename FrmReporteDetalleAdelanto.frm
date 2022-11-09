VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDetalleAdelanto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   6240
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TxtSaldo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4155
      MaxLength       =   80
      TabIndex        =   22
      Top             =   2000
      Width           =   1575
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Informe"
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
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
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CheckBox ChkAlmacen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Almacen"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   510
      Width           =   1095
   End
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   840
      MaxLength       =   80
      TabIndex        =   4
      Top             =   1260
      Width           =   735
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   3
      Top             =   1260
      Width           =   3975
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1620
      Width           =   3975
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   1
      Top             =   2000
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   2685
      TabIndex        =   8
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin MSComCtl2.DTPicker DTPDesde 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3240
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
      Format          =   17104897
      CurrentDate     =   39788
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdAdelantos 
      Height          =   1575
      Left            =   435
      TabIndex        =   12
      Top             =   4080
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
      Height          =   5895
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   7215
      Begin VB.CommandButton CmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Adelantos:"
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   6975
         Begin VB.Label LblAdelantos 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4680
            TabIndex        =   18
            Top             =   1845
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Salidas de Dinero:"
         ForeColor       =   &H000000FF&
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   3480
         Width           =   6975
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSalidas 
            Height          =   1575
            Left            =   120
            TabIndex        =   15
            Top             =   240
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
         Begin VB.Label LblSalidas 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   16
            Top             =   1920
            Width           =   2025
         End
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
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
         Format          =   17104897
         CurrentDate     =   39788
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
         TabIndex        =   21
         Top             =   840
         Width           =   135
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
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
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
      Left            =   3450
      TabIndex        =   23
      Top             =   2040
      Width           =   615
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
      Left            =   720
      TabIndex        =   6
      Top             =   1620
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
      Left            =   885
      TabIndex        =   5
      Top             =   1980
      Width           =   555
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
      Left            =   855
      TabIndex        =   0
      Top             =   960
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   840
      Top             =   315
      Width           =   4815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmDetalleAdelanto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim RstSalidas As New ADODB.Recordset
Dim TotalAdelanto As Double
Dim TotalSalidas As Double


Private Sub Check1_Click()

End Sub

Private Sub ChkSalidaDinero_Click()
Dim Anulado As String
Anulado = "F"
If Me.ChkSalidaDinero.Value = 1 Then
   ' Me.HfdAdelanto
    strCadena = "SELECT IdAdelanto as NUM , Fecha,Cliente, MontoAdelantado " & _
    "FROM DetalleAdelantos " & _
    "WHERE DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "'  ORDER BY 2 ASC"
    Call LLenaGrilla
    strCadena = "SELECT DocumentoVenta.dEmisionVenta as Fecha,(Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta)Numero, " & _
            "DocumentoVenta.cPersona , DocumentoVenta.nTotalVenta FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND " & _
            "DocumentoVenta.idFormaPago='" & KEY_ADELANTO & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "'"
    Call LlenarGrillaSalida(strCadena)
Else
    Me.HfgSalidas.Visible = True
    Me.CmdImprimir.Visible = False
End If
End Sub


Private Sub CmdFiltrar_Click()
Dim Anulado As String
Anulado = "F"
strCadena = "SELECT IdAdelanto as NUM , Fecha,Cliente, MontoAdelantado " & _
    "FROM DetalleAdelantos " & _
    "WHERE DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "' AND DetalleAdelantos.Fecha>='" & CVDate(Me.DtpDesde.Value) & "' AND DetalleAdelantos.Fecha<='" & CVDate(Me.DtpHasta.Value) & "' ORDER BY 2 ASC"
    Call LLenaGrilla
    
    strCadena = "SELECT DocumentoVenta.dEmisionVenta as Fecha,(Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta)Numero, " & _
            "DocumentoVenta.cPersona , DocumentoVenta.nTotalVenta FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND " & _
            "DocumentoVenta.idFormaPago='" & KEY_ADELANTO & "' AND DocumentoVenta.dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND DocumentoVenta.dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "'"
        
     Call LlenarGrillaSalida(strCadena)
     Me.CmdImprimir.Enabled = True
        
End Sub

Private Sub LlenarGrillaSalida(ByVal Cadena As String)
    Set RstSalidas = Nothing
    RstSalidas.Open Cadena, CnBd, adOpenKeyset, adLockOptimistic
         If RstSalidas.RecordCount > 0 Then
        Me.CmdImprimir.Visible = True
        Me.HfgSalidas.Visible = True
        Me.HfgSalidas.Clear
        Set Me.HfgSalidas.Recordset = RstSalidas
        HfgSalidas.ColWidth(0) = 1050
        HfgSalidas.ColWidth(1) = 3000
        HfgSalidas.ColWidth(2) = 1000
        HfgSalidas.ColWidth(3) = 1500
        Call DarFormatoFecha(HfgSalidas, 0)
        Call DarFormato(HfgSalidas, 3)
        RstSalidas.MoveFirst
        For i = 0 To RstSalidas.RecordCount - 1
             Monto = Monto + RstSalidas(3)
             RstSalidas.MoveNext
        Next i
        Me.LblSalidas.Caption = "S/." + Space(2) + Format(Str(Monto), "#,##0.00")
        TotalSalidas = Monto
        Me.TxtSaldo.Text = Str(Format(TotalAdelanto - TotalSalidas, "#,##0.000"))
        
     End If
Set RstSalidas = Nothing
End Sub
Private Sub CmdImprimir_Click()
    Dim Anulado As String
    Dim TotalP As String
    Dim total As Double
    Dim fecha As String
    Dim Documento As String
    Dim Numero As Integer
    Dim usuario As String
    Dim totalletras As String
    Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 10
    Anulado = "F"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "COD      :" + Space(1); Me.TxtCodCliente.Text
    Printer.Print Tab(4); "CLIENTE  :" + Space(1); Mid(Me.TxtCliente.Text + Space(80), 1, 40) & Space(1) & CVDate(Date)
    Printer.Print ""
    Printer.Print Tab(4); "DIRECCIÓN:" + Space(1); Mid(Me.TxtDireccion.Text + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(4); "REPORTE AL:" + Space(1) + Str(Date) + Space(2) + "A LAS:" + Space(1) + Str(Time)
    Printer.Print "----------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(4); "ADELANTOS:"
    
    
    strCadena = "SELECT IdAdelanto as NUM ,Fecha,Cliente, MontoAdelantado FROM DetalleAdelantos " & _
    "WHERE (DetalleAdelantos.IdCliente='" & Trim(Me.TxtCodCliente.Text) & "' AND DetalleAdelantos.Fecha>='" & CVDate(Me.DtpDesde.Value) & "' AND DetalleAdelantos.Fecha<='" & CVDate(Me.DtpHasta.Value) & "') ORDER BY 1 ASC"
    Call ConfiguraRst(strCadena)
    
    strCadena = "SELECT DocumentoVenta.dEmisionVenta as Fecha,(Comprobantes.doc_abrev +'-'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta)Numero, " & _
            "DocumentoVenta.cPersona , DocumentoVenta.nTotalVenta FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND " & _
            "DocumentoVenta.idFormaPago='" & KEY_ADELANTO & "' AND DocumentoVenta.dEmisionVenta>='" & CVDate(Me.DtpDesde.Value) & "' AND DocumentoVenta.dEmisionVenta<='" & CVDate(Me.DtpHasta.Value) & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "'"
            RstSalidas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
      
    
    rst.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    total = 0
            For j = 0 To rst.RecordCount - 1
                Numero = Mid(Str(rst(0)) + Space(10), 1, 5)
                fecha = CVDate(rst(1))
                total = total + rst(3)
                TotalP = Mid(Format(Str(rst(3)), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(1); Numero & Space(2) & fecha & Space(25) & TotalP
                Printer.CurrentY = Printer.CurrentY + 0.4
                rst.MoveNext
            Next j
            inc = 0.5
            
            Printer.Print Tab(40); "==================="
            Printer.Print Tab(40); Format(Str(total), "#,##0.00")
            
            Printer.Print ""
    
    total = Str(total)
    totalletras = UCase(EnLetras(Format(Str(total), "#,##0.00")))
    Set rst = Nothing
    '---- fin totales
    
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(5); Mid(totalletras + Space(100), 1, 100)
    Printer.Print "----------------------------------------------------------------------"
    Printer.Print Tab(4); "SALIDAS:"
    Call ConfiguraRst(strCadena)
    If RstSalidas.EOF = False Then
        RstSalidas.MoveFirst
    Else
        Set RstSalidas = Nothing
        Exit Sub
    End If
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    total = 0
            For j = 0 To RstSalidas.RecordCount - 1
                fecha = CVDate(rst(0))
                Documento = Mid(RstSalidas(1) + Space(30), 1, 30)
                total = total + RstSalidas(3)
                TotalP = Mid(Format(Str(RstSalidas(3)), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(1); fecha & Space(3) & Documento & Space(3) & TotalP
                Printer.CurrentY = Printer.CurrentY + 0.4
                RstSalidas.MoveNext
            Next j
            inc = 0.5
            
            Printer.Print Tab(40); "==================="
            Printer.Print Tab(47); Format(Str(total), "#,##0.00")
                        
    total = Str(total)
    totalletras = UCase(EnLetras(Format(Str(total), "#,##0.00")))
    
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(5); Mid(totalletras + Space(100), 1, 100)
    Printer.Print "----------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print Tab(10); "SALDO:" + Space(2); Me.TxtSaldo.Text
    Printer.Print "----------------------------------------------------------------------"
    Set rst = Nothing
    Set RstSalidas = Nothing
    Printer.EndDoc
    
    Exit Sub
    
    
     

    
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodCliente)
End If
End Sub



Private Sub Form_Activate()
Me.DtcAlmacen.SetFocus
Me.CmdImprimir.Enabled = False
End Sub


Private Sub LLenaGrilla()
Dim Monto As Double
Dim i As Integer
Monto = 0
     Call ConfiguraRst(strCadena)
       Me.HfdAdelantos.Clear
     If rst.RecordCount > 0 Then
        Me.CmdImprimir.Visible = True
        Me.HfdAdelantos.Visible = True
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
        Me.LblAdelantos.Caption = "S/." + Space(2) + Format(Str(Monto), "#,##0.00")
        TotalAdelanto = Monto
     End If
Set rst = Nothing
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
doc_Tienda = "V"
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.ChkAlmacen.Value = 1
   Set rst = Nothing
  Me.DtpDesde.Value = CVDate(Date)
  Me.DtpHasta.Value = CVDate(Date)
  
    
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.Key
    Case KEY_NEW
      'Call Nuevo
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


Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtCodCliente.Text = FormatosCeros(Me.TxtCodCliente.Text, 5)
    strCadena = "SELECT cPersona,NombrePersona,sRazonSocial,sDireccionCliente1,Per_Ruc,Observacion,MontoAdelantado FROM " & _
    " Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Trim(rst(6)) = "J" Then
            Me.TxtCliente.Text = rst(2)
        Else
            Me.TxtCliente.Text = rst(1)
        End If
        
        Me.TxtDireccion.Text = rst(3)
        Me.TxtRuc.Text = rst(4)
        Me.TxtSaldo.Text = "Pulse check" 'Format(Rst(7), "#,##0.00")
        
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









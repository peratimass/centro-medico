VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmTempExportExcel 
   BorderStyle     =   0  'None
   Caption         =   "Exportar a Excel"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgGuardar 
      Left            =   480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraOpciones 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   7455
      Begin VB.TextBox txtidperiodo 
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DtcSucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkSucursal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "POR SUCURSAL:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   3960
         TabIndex        =   18
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optdetallado 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "DETALLADO"
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
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optresumen 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "RESUMIDO"
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
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VitekeySoft.TextBoxPlus txtNRegs 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkNRegs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "DIVIDIR CADA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Label lblNRegs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REGISTROS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.PictureBox FraColumnas 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Frame fraCustom 
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4215
         Begin VB.CommandButton cmdDelAll 
            Caption         =   "<<"
            Height          =   255
            Left            =   1920
            TabIndex        =   10
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton cmdDelOne 
            Caption         =   "<"
            Height          =   255
            Left            =   1920
            TabIndex        =   9
            Top             =   1200
            Width           =   375
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   ">>"
            Height          =   255
            Left            =   1920
            TabIndex        =   8
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdAddOne 
            Caption         =   ">"
            Height          =   255
            Left            =   1920
            TabIndex        =   7
            Top             =   360
            Width           =   375
         End
         Begin VB.ListBox lstDestino 
            Height          =   1815
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.ListBox lstOrigen 
            Height          =   1815
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.OptionButton optSoloValores 
         Caption         =   "Solo columnas con valores"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   3015
      End
      Begin VB.OptionButton optCustom 
         Caption         =   "Personalizado"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "TODOS"
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
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000C0C0&
         TabIndex        =   1
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
   End
   Begin MSComctlLib.ProgressBar pbrAvance 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmTempExportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public IdEmpPeriodo As Long
Public FormatNom As String
Public PeriodoNom As String
Public nomArchivo As String
Public TipoReg As Integer
Public IdCliente As Long

Dim UltColumna As Long
Dim FinFilaCab As Long
Dim ColNomEntidad As Long
Dim ColNumero As Long

Dim PiePagina As String

Dim FlagTotal As Boolean
Dim FlagFecVec As Boolean
Dim FlagAnDua As Boolean

Dim FlagExport As Boolean
Dim FlagAfecto As Boolean
Dim FlagExonerado As Boolean
Dim FlagInafecto As Boolean
Dim FlagISC As Boolean
Dim FlagIGV As Boolean
Dim FlagBIvap As Boolean
Dim FlagIvap As Boolean

Dim FlagMGrv1 As Boolean
Dim FlagMIgv1 As Boolean
Dim FlagMGrv2 As Boolean
Dim FlagMIgv2 As Boolean
Dim FlagMGrv3 As Boolean
Dim FlagMIgv3 As Boolean
Dim FlagNoGrav As Boolean
Dim FlagNoDom As Boolean
Dim FlagDetracNum As Boolean
Dim FlagDetracFec As Boolean
Dim FlagNaturaleza As Boolean
Dim FlagCcostos1 As Boolean
Dim FlagCcostos2 As Boolean
Dim FlagCcostos3 As Boolean

Dim FlagDetalle As Boolean
Dim FlagCuenta As Boolean
Dim FlagCCostos As Boolean

Dim FlagOtros As Boolean
Dim FlagPerRet As Boolean
Dim FlagRetencion As Boolean
Dim FlagPercepcion As Boolean
Dim FlagTipoCambio As Boolean
Dim FlagRFecha As Boolean
Dim FlagRTipo As Boolean
Dim FlagRSerie As Boolean
Dim FlagRNumero As Boolean
Dim FlagTipoOpe As Boolean
Dim FlagCantIng As Boolean
Dim FlagCostoIng As Boolean
Dim FlagSaldoIng As Boolean
Dim FlagCantSal As Boolean
Dim FlagCostoSal As Boolean
Dim FlagSaldoSalida As Boolean

Dim FlagCantSaldo As Boolean
Dim FlagCostoPomedio As Boolean
Dim FlagSaldoFinal As Boolean


Private Sub chkSucursal_Click()
If Me.chkSucursal.Value = 1 Then
   Me.DtcSucursal.Visible = True
Else
    Me.DtcSucursal.Visible = False
End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_tipoentidad='0' ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcSucursal)
    Me.DtcSucursal.BoundText = KEY_ALM
    
    If FrmRegistroVentas.Procedencia = nuevo Then
        Me.Top = FrmRegistroVentas.Top + 500
        Me.Caption = "RUC:" & KEY_RUC + Space(5) + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 2)
    End If
    
    If FrmRegistroCompras.Procedencia = nuevo Then
        Me.Top = FrmRegistroCompras.Top + 500
        Me.Caption = "RUC:" & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 0) + Space(5) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2)
    End If
    If FrmKardexdeProductos.Procedencia = nuevo Then
        Me.Top = FrmKardexdeProductos.Top + 500
        Me.Caption = "RUC:" & KEY_RUC + Space(5) + KEY_EMPRESA
    End If
    
    dlgGuardar.InitDir = "%HOMEDRIVE%"
    dlgGuardar.Filter = "Archivo de Excel (*.xlsx)|*.xlsx"
    dlgGuardar.FilterIndex = 1
    dlgGuardar.DialogTitle = "Guardar reporte como..."
    'dlgGuardar.FileName = NomArchivo
    
    If FrmRegistroVentas.Procedencia = nuevo Then
        dlgGuardar.FileName = "RegVentas" + Space(1) + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1) + "-" + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3) + "-" + KEY_RUC
    End If
    
    If FrmKardexdeProductos.Procedencia = nuevo Then
        dlgGuardar.FileName = "KardexValorizado" + Space(1) + FrmKardexdeProductos.DtcProducto.BoundText + "-" + KEY_RUC
    End If
    
    If FrmRegistroCompras.Procedencia = nuevo Then
        dlgGuardar.FileName = "RegCompras" + Space(1) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1) + "-" + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3) + "-" + KEY_RUC
    End If
    If FrmDaotCompra.Procedencia = nuevo Then
        dlgGuardar.FileName = "DaotCompras" + Space(1) + KEY_RUC
    End If
    
    dlgGuardar.Flags = _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNPathMustExist
    dlgGuardar.CancelError = True
    
    'optSoloValores.Value = True
    Me.optTodos.Value = True
    If FrmRegistroVentas.Procedencia = nuevo Then
        TipoReg = 1
    End If
    If FrmRegistroCompras.Procedencia = nuevo Then
        TipoReg = 0
    End If
     
     If FrmDaotCompra.Procedencia = nuevo Then
        TipoReg = 3
     End If
     
     If FrmKardexdeProductos.Procedencia = nuevo Then
        TipoReg = 10
     End If
     
    
    Select Case TipoReg
        Case 1 'Ventas
            Call addELista(lstOrigen, 2, "Fec.Vec.")
            Call addELista(lstOrigen, 9, "Exportaci�n")
            Call addELista(lstOrigen, 10, "Afecto")
            Call addELista(lstOrigen, 11, "Exonerado")
            Call addELista(lstOrigen, 12, "Inafecto")
            Call addELista(lstOrigen, 13, "ISC")
            Call addELista(lstOrigen, 14, "IGV")
            Call addELista(lstOrigen, 15, "Base IVAP")
            Call addELista(lstOrigen, 16, "IVAP")
            Call addELista(lstOrigen, 17, "Retenci�n")
            Call addELista(lstOrigen, 18, "Otros")
            Call addELista(lstOrigen, 19, "Total")
            Call addELista(lstOrigen, 20, "Tipo Cambio")
            Call addELista(lstOrigen, 21, "RD Fecha")
            Call addELista(lstOrigen, 22, "RD Tipo")
            Call addELista(lstOrigen, 23, "RD Serie")
            Call addELista(lstOrigen, 24, "RD Numero")
            
            PiePagina = "VENTAS"
        
        Case 0 'Compras
            Call addELista(lstOrigen, 1, "Fec.Vec.")
            Call addELista(lstOrigen, 2, "A�o DUA")
            Call addELista(lstOrigen, 10, "M.Gravado1")
            Call addELista(lstOrigen, 11, "IGV 1")
            Call addELista(lstOrigen, 12, "M.Gravado2")
            Call addELista(lstOrigen, 13, "IGV 2")
            Call addELista(lstOrigen, 14, "M.Gravado3")
            Call addELista(lstOrigen, 15, "IGV 3")
            Call addELista(lstOrigen, 16, "No Gravado")
            Call addELista(lstOrigen, 17, "ISC")
            Call addELista(lstOrigen, 18, "Percepci�n")
            Call addELista(lstOrigen, 19, "Otros")
            Call addELista(lstOrigen, 20, "Total")
            Call addELista(lstOrigen, 21, "N.Comp. No Domiciliado")
            Call addELista(lstOrigen, 22, "N.Const.Detracci�n")
            Call addELista(lstOrigen, 23, "Fec.Const.Detracci�n")
            Call addELista(lstOrigen, 24, "Tipo Cambio")
            Call addELista(lstOrigen, 25, "RD Fecha")
            Call addELista(lstOrigen, 26, "RD Tipo")
            Call addELista(lstOrigen, 27, "RD Serie")
            Call addELista(lstOrigen, 28, "RD Numero")
            
            PiePagina = "COMPRAS"
            
        Case 6 'Gastos
            Call addELista(lstOrigen, 2, "Fec.Vec.")
            Call addELista(lstOrigen, 2, "Detalle")
            Call addELista(lstOrigen, 10, "Inafecto")
            Call addELista(lstOrigen, 12, "Afecto")
            Call addELista(lstOrigen, 14, "IGV")
            Call addELista(lstOrigen, 18, "Otros")
            Call addELista(lstOrigen, 19, "Total")
            Call addELista(lstOrigen, 20, "Tipo Cambio")
            Call addELista(lstOrigen, 21, "RD Fecha")
            Call addELista(lstOrigen, 22, "RD Tipo")
            Call addELista(lstOrigen, 23, "RD Serie")
            Call addELista(lstOrigen, 24, "RD Numero")
            Call addELista(lstOrigen, 25, "Cuenta")
            Call addELista(lstOrigen, 26, "Centro Costos")
            
            PiePagina = "GASTOS"
         Case 3 'Daot
            Call addELista(lstOrigen, 2, "Fecha")
            Call addELista(lstOrigen, 2, "RUC")
            Call addELista(lstOrigen, 10, "RAZON SOCIAL")
            Call addELista(lstOrigen, 12, "VALOR VENTA")
            Call addELista(lstOrigen, 14, "IGV")
            Call addELista(lstOrigen, 18, "TOTAL")
            PiePagina = "DAOT COMPRA"
        Case 10 'Kardex
            Call addELista(lstOrigen, 2, "Fecha")
            Call addELista(lstOrigen, 9, "Tipo")
            Call addELista(lstOrigen, 10, "Serie")
            Call addELista(lstOrigen, 11, "Numero")
            Call addELista(lstOrigen, 12, "Tipo Operacion")
            Call addELista(lstOrigen, 13, "Cantidad")
            Call addELista(lstOrigen, 14, "Costo U.")
            Call addELista(lstOrigen, 15, "Total")
            Call addELista(lstOrigen, 16, "Cantidad")
            Call addELista(lstOrigen, 17, "Costo U.")
            Call addELista(lstOrigen, 18, "Cantidad")
            Call addELista(lstOrigen, 19, "Costo U")
            Call addELista(lstOrigen, 20, "Total")
            
            
            PiePagina = "INVENTARIO VALIRZADO"
    End Select
     If FrmRegistroVentas.Procedencia = nuevo Then
        PiePagina = PiePagina & " " & FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 2) & "    " & BDBuscarCampo("persona", "dni", "dni", str(FrmRegistroVentas.txtRuc.Text))
    End If
    If FrmRegistroCompras.Procedencia = nuevo Then
        PiePagina = PiePagina & " " & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2) & "    " & BDBuscarCampo("persona", "dni", "dni", str(FrmRegistroCompras.txtRuc.Text))
    End If
    If FrmDaotCompra.Procedencia = nuevo Then
        PiePagina = PiePagina & " " & FrmDaotCompra.DtpInicio.Value & "    " & FrmDaotCompra.DtpFinal.Value
    End If
    If FrmKardexdeProductos.Procedencia = nuevo Then
        PiePagina = PiePagina
    End If
    
End Sub

Private Sub chkNRegs_Click()
    If chkNRegs.Value = 0 Then
        txtNRegs.Enabled = False
    Else
        txtNRegs.Enabled = True
    End If
End Sub

Private Sub cmdAddOne_Click()
    If lstOrigen.ListIndex >= 0 Then
        Call addELista(lstDestino, lstOrigen.ItemData(lstOrigen.ListIndex), lstOrigen.Text)
        lstOrigen.RemoveItem (lstOrigen.ListIndex)
        If lstOrigen.ListCount > 0 Then
            lstOrigen.Selected(0) = True
        End If
    End If
End Sub

Private Sub cmdAddAll_Click()
    Dim i As Integer
    For i = 0 To lstOrigen.ListCount - 1
        lstOrigen.Selected(i) = True
        Call addELista(lstDestino, lstOrigen.ItemData(i), lstOrigen.Text)
    Next i
    lstOrigen.Clear
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
    If FrmRegistroVentas.Procedencia = nuevo Then
        FrmRegistroVentas.Procedencia = Neutro
    End If
    If FrmRegistroCompras.Procedencia = nuevo Then
        FrmRegistroCompras.Procedencia = Neutro
    End If
End Sub

Private Sub cmdDelOne_Click()
    If lstDestino.ListIndex >= 0 Then
        Call addELista(lstOrigen, lstDestino.ItemData(lstDestino.ListIndex), lstDestino.Text)
        lstDestino.RemoveItem (lstDestino.ListIndex)
        If lstDestino.ListCount > 0 Then
            lstDestino.Selected(0) = True
        End If
    End If
End Sub


Private Sub cmdDelAll_Click()
    Dim i As Integer
    For i = 0 To lstDestino.ListCount - 1
        lstDestino.Selected(i) = True
        Call addELista(lstOrigen, lstDestino.ItemData(i), lstDestino.Text)
    Next i
    lstDestino.Clear
End Sub


Private Function SelArchivo() As String
    On Error Resume Next
    
    dlgGuardar.ShowSave
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Exit Function
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        MsgBox "Error " & Format$(Err.Number) & _
            " seleccionando ruta." & vbCrLf & _
            Err.Description
        Exit Function
    End If
    SelArchivo = dlgGuardar.FileName
End Function
Private Sub cmdExport_Click()
    
    Dim RutaArchivo As String
    Dim in_producto As String
    Dim in_kardex As Boolean
    
    
    
    in_kardex = False
    
    If FrmRegistroVentas.Procedencia = nuevo Then
   
       If Me.chkSucursal.Value = 1 Then
          in_alm = Me.DtcSucursal.BoundText
       Else
          in_alm = ""
       End If
       
        If Me.optresumen.Value = True Then
           Call exportar_ventas(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 0), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3), in_alm)
        End If
        If Me.optdetallado.Value = True Then
           Call exportar_ventas_detallado(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 0), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3), in_alm)
        End If
    End If
    
    
    
    
    
    
    
    
    If FrmKardexdeProductos.Procedencia = nuevo Then
        If Me.chkSucursal.Value = 1 Then
          in_alm = Me.DtcSucursal.BoundText
       Else
          in_alm = ""
       End If
       in_producto = FrmKardexdeProductos.txtcodigoprod.Text
       
        in_kardex = True
    End If
    
    
   
 
    
    RutaArchivo = SelArchivo
    If RutaArchivo = vbNullString Then Exit Sub
    
    FraColumnas.Enabled = False
    fraOpciones.Enabled = False
    cmdExport.Enabled = False
    cmdCancelar.Enabled = False
    If in_kardex = True Then
        
    If Exportar_Excel_vv(RutaArchivo) Then
        FrmRegistroVentas.Procedencia = Neutro
        MsgBox "La exportaci�n concluy� satisfactoriamente"
        Call AbreArchivo(RutaArchivo)
    End If
    Else
        If Exportar_Excel(RutaArchivo) Then
            FrmRegistroVentas.Procedencia = Neutro
            MsgBox "La exportaci�n concluy� satisfactoriamente"
            Call AbreArchivo(RutaArchivo)
    End If
    End If
    
    
    
    ResetFlags
    FraColumnas.Enabled = True
    fraOpciones.Enabled = True
    cmdExport.Enabled = True
    cmdCancelar.Enabled = True
End Sub
Public Sub exportar_ventas(ByVal in_ruc As String, ByVal in_mes As String, ByVal in_anio As String, ByVal in_alm As String)
    
    
    Dim bol_ini As String
    Dim bol_fin As String
    Dim i As Double
    Dim fecha As Date
    Dim Total As Double
    Dim Total_Registros As Double
    
    strCadena = "DELETE FROM registroventassunat WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' "
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and  ruc='" & KEY_RUC & " ' and month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc IN('0001','0007','0008')  ORDER BY fecha_emision,id_doc,serie,numero ASC"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
        fecha = rst("fecha_emision")
        Me.pbrAvance.Min = 0
        Me.pbrAvance.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
        If Len(Trim(rst("id_cliente"))) = 11 Then
            tdoc = 6
        Else
            tdoc = 0
        End If
        
        
        
        
        
        If rst("anulado") = "si" Then
            n_total = 0
            n_subtotal = 0
            n_igv = 0
            n_exonerado = 0
        Else
            If rst("id_moneda") = "00001" Then
                n_total = rst("total")
                If rst("igv") > 0 Then
                    n_subtotal = n_total / (1 + KEY_IGV)
                    n_igv = n_total - n_subtotal
                    n_exonerado = rst("exonerado")
                Else
                    n_subtotal = 0
                    n_igv = 0
                    n_exonerado = rst("exonerado")
                End If
                
            Else
                
                If rst("igv") > 0 Then
                    n_total = rst("total") * rst("tc")
                    n_subtotal = rst("valor_venta") * rst("tc")
                    n_igv = rst("igv") * rst("tc")
                    n_exonerado = rst("exonerado") * rst("tc")
                Else
                    n_total = rst("total") * rst("tc")
                    n_subtotal = 0
                    n_igv = 0
                    n_exonerado = rst("exonerado") * rst("tc")
                End If
                
                
            End If
        End If
        If rst("id_doc") = "0007" Then
                
                If rst("igv") > 0 Then
                    n_total = n_total * -1
                    n_subtotal = (n_total / (1 + KEY_IGV))
                    n_igv = (n_total - n_subtotal)
                    n_exonerado = n_exonerado * -1
                Else
                    n_total = n_total * -1
                    n_subtotal = 0
                    n_igv = 0
                    n_exonerado = n_exonerado * -1
                End If
                
        End If
               
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,retencion,tc,fecha_F,doc_codF,serieF,numeroF,dni_save,anulado)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY/mm/dd") & "','" & in_mes & "','" & in_anio & "','" & Trim(formato_item(Val(rst("id_doc")), 2)) & "'," & _
        "'" & rst("serie") & "','" & rst("numero") & "','" & tdoc & "','" & rst("id_cliente") & "','" & Trim(Replace(rst("ncliente"), "'", "")) & "','" & n_subtotal & "','" & n_exonerado & "','" & n_igv & "'," & _
        "'" & n_total & "','" & rst("retencion") & "','" & rst("tc") & "','" & Format(rst("fecha_fact"), "YYYY-mm-dd") & "','" & rst("id_doc_fact") & "','" & rst("serie_fact") & "','" & rst("numero_fact") & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
        
        Me.pbrAvance.Value = i
        rst.MoveNext
Next i
End If


strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'   and  ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0003' AND anulado='no' ORDER BY serie ASC"

Call ConfiguraRstT(strCadena)


strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and  ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0003'  ORDER BY fecha_emision ASC"

Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
    Me.pbrAvance.Min = 0
    Me.pbrAvance.Max = rst.RecordCount
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
        
        strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and  fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0003'"
        
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                If rstTemporal(2) > 0 Then
                    afecto = rstTemporal(0)
                Else
                    afecto = 0
                End If
                
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
                afecto = 0
                igv = 0
            Else
                
                Total = rstTemporal(3)
                If KEY_CON_IGV = "si" Then
                    afecto = Total / (1 + KEY_IGV)
                    igv = Total - afecto
                Else
                    afecto = 0
                    igv = 0
                End If
                
            End If
        
            strCadena = "SELECT serie,numero FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and  fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0003'  ORDER BY numero ASC"
        
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & in_mes & "','" & in_anio & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
        
    End If
avanzar:

    
    rstT.MoveNext
Next j
    Me.pbrAvance.Value = i
    rst.MoveNext
Next i
End If


strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0012' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)

strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0012' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
    Me.pbrAvance.Min = 0
    Me.pbrAvance.Max = rst.RecordCount
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
    strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0012'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                afecto = rstTemporal(0)
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
            Else
                Total = rstTemporal(3)
            End If
        strCadena = "SELECT serie,numero FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND month(fecha_emision)='" & Val(in_mes) & "' and  id_anio='" & in_anio & "' AND id_doc='0012'  ORDER BY numero ASC"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar1
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save,anulado)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & in_mes & "','" & in_anio & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
         
    End If
avanzar1:
    rstT.MoveNext
Next j
    Me.pbrAvance.Value = i
    rst.MoveNext
Next i
End If


End Sub

Public Sub exportar_ventas_detallado(ByVal in_ruc As String, ByVal in_mes As String, ByVal in_anio As String, ByVal in_alm As String)
    
    
    Dim bol_ini As String
    Dim bol_fin As String
    Dim i As Double
    Dim fecha As Date
    Dim Total As Double
    Dim Total_Registros As Double
    
    strCadena = "SELECT igv FROM entidad_parametros where cod_unico='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    KEY_CON_IGV = rst("igv")
    
    strCadena = "DELETE FROM registroventassunat WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' "
    CnBd.Execute (strCadena)
    strCadena = "SELECT * FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & " ' and month(fecha_emision)='" & Val(in_mes) & "' and id_anio='" & in_anio & "' AND id_doc IN('0001','0003','0007','0008')  ORDER BY fecha_emision,id_doc,serie,numero ASC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
        fecha = rst("fecha_emision")
        Me.pbrAvance.Min = 0
        Me.pbrAvance.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
        
    
        If Len(Trim(rst("id_cliente"))) = 11 Then
            tdoc = 6
        Else
            tdoc = 0
        End If
        
        If rst("anulado") = "si" Then
            n_total = 0
            n_subtotal = 0
            n_igv = 0
            n_exonerado = 0
        Else
            
            
            If KEY_PAIS <> KEY_PERU Then
                If rst("id_moneda") = "00002" Then
                    n_total = rst("total")
                If KEY_CON_IGV = "si" Then
                   n_subtotal = n_total / 1.18
                   n_igv = n_total - n_subtotal
                   n_exonerado = rst("exonerado")
                Else
                   n_subtotal = n_total
                   n_igv = n_total - n_subtotal
                   n_exonerado = rst("exonerado")
                   n_subtotal = 0
                End If
                
                
                
            Else
                n_total = rst("total") * rst("tc")
                n_subtotal = rst("valor_venta") * rst("tc")
                n_igv = rst("igv") * rst("tc")
                n_exonerado = rst("exonerado") * rst("tc")
            End If
            
            
            
            Else
                If rst("id_moneda") = "00001" Then
                n_total = rst("total")
                If KEY_CON_IGV = "si" Then
                   n_subtotal = n_total / 1.18
                   n_igv = n_total - n_subtotal
                   n_exonerado = rst("exonerado")
                Else
                   n_subtotal = n_total
                   n_igv = n_total - n_subtotal
                   n_exonerado = rst("exonerado")
                   n_subtotal = 0
                End If
                
                
                
            Else
                n_total = rst("total") * rst("tc")
                n_subtotal = rst("valor_venta") * rst("tc")
                n_igv = rst("igv") * rst("tc")
                n_exonerado = rst("exonerado") * rst("tc")
            End If
            
            End If
            
            
        
        
        
        
        End If
        If rst("id_doc") = "0007" Then
                
                If KEY_CON_IGV = "si" Then
                    n_total = n_total * -1
                    n_subtotal = (n_total / 1.18)
                    n_igv = (n_total - n_subtotal)
                    n_exonerado = n_exonerado * -1
                Else
                    n_total = n_total * -1
                    n_subtotal = 0
                    n_igv = 0
                    n_exonerado = n_total
                End If
                
        End If
               
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,retencion,tc,fecha_F,doc_codF,serieF,numeroF,dni_save,anulado)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY/mm/dd") & "','" & in_mes & "','" & in_anio & "','" & Trim(formato_item(Val(rst("id_doc")), 2)) & "'," & _
        "'" & rst("serie") & "','" & rst("numero") & "','" & tdoc & "','" & rst("id_cliente") & "','" & Trim(Replace(rst("ncliente"), "'", "")) & "','" & n_subtotal & "','" & n_exonerado & "','" & n_igv & "'," & _
        "'" & n_total & "','" & rst("retencion") & "','" & rst("tc") & "','" & Format(rst("fecha_fact"), "YYYY-mm-dd") & "','" & rst("id_doc_fact") & "','" & rst("serie_fact") & "','" & rst("numero_fact") & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
        
        Me.pbrAvance.Value = i
        rst.MoveNext
Next i
End If


Exit Sub
strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0003' AND anulado='no' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)

strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0003' AND anulado='no' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
    Me.pbrAvance.Min = 0
    Me.pbrAvance.Max = rst.RecordCount
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
        strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0003' and anulado='no'"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                afecto = rstTemporal(0)
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
                afecto = 0
                igv = 0
            Else
                Total = rstTemporal(3)
                afecto = Total / 1.18
                igv = Total - afecto
            End If
        strCadena = "SELECT serie,numero FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0003' and anulado='no'  ORDER BY numero ASC"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & in_mes & "','" & in_anio & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
        
    End If
avanzar:

    
    rstT.MoveNext
Next j
    Me.pbrAvance.Value = i
    rst.MoveNext
Next i
End If

strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0012' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)

strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0012' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
    Me.pbrAvance.Min = 0
    Me.pbrAvance.Max = rst.RecordCount
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
    strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0012'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                afecto = rstTemporal(0)
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
            Else
                Total = rstTemporal(3)
            End If
        strCadena = "SELECT serie,numero FROM movimiento_venta WHERE id_alm LIKE '%" & in_alm & "%'  and fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & KEY_RUC & "' AND id_mes='" & in_mes & "' AND id_anio='" & in_anio & "' AND id_doc='0012'  ORDER BY numero ASC"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar1
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save,anulado)VALUES " & _
        "('" & in_ruc & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & in_mes & "','" & in_anio & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
        
    End If
avanzar1:
    rstT.MoveNext
Next j
    Me.pbrAvance.Value = i
    rst.MoveNext
Next i
End If


End Sub
Private Function Exportar_Excel(sOutputPath As String) As Boolean
  
   ' On Error GoTo error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim iSheetsPerBook As Integer
    
    Dim Fila        As Long
    Dim columna     As Long
    
    Dim IniSeccionData As Long
    
    Dim FactorZoom As Integer
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    
    iSheetsPerBook = o_Excel.SheetsInNewWorkbook
    o_Excel.SheetsInNewWorkbook = 1
    
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
    o_Hoja.name = "Reporte"
    
    o_Excel.SheetsInNewWorkbook = iSheetsPerBook
    
    Fila = 1
    columna = 1
    
    ' -- Grabar Datos de Cabecera
    o_Hoja.Cells(Fila, columna).Value = FormatNom
    o_Hoja.Cells(Fila, columna).Font.Size = 13
    o_Hoja.Cells(Fila, columna).Font.Bold = True
    
   If FrmRegistroVentas.Procedencia = nuevo Then
     If Me.chkSucursal.Value = 1 Then
        in_sucursal = Me.DtcSucursal.Text
    Else
        in_sucursal = ""
     End If
     d_periodo = "PERIODO" & Mid(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 2), 18, 15) + Space(1) + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3) & Space(1) + in_sucursal
     d_anio = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)
     d_ruc = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 0)
     o_Hoja.Cells(Fila + 1, columna).Value = "FORMATO 14.1: REGISTRO DE VENTAS E INGRESOS"
  End If
     
  If FrmRegistroCompras.Procedencia = nuevo Then
     d_periodo = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2)
     d_anio = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
     d_ruc = FrmRegistroCompras.txtRuc.Text
     o_Hoja.Cells(Fila + 1, columna).Value = "FORMATO 8.4: REGISTRO DE COMPRAS"
   End If
  
  If FrmDaotCompra.Procedencia = nuevo Then
     d_periodo = str(FrmDaotCompra.DtpInicio.Value) + "-" + str(FrmDaotCompra.DtpFinal.Value)
     d_anio = Year(KEY_FECHA)
     d_ruc = KEY_RUC
  End If
  
  If FrmKardexdeProductos.Procedencia = nuevo Then
     If Me.chkSucursal.Value = 1 Then
        in_sucursal = Me.DtcSucursal.Text
    Else
        in_sucursal = ""
     End If
     d_periodo = "PERIODO:"
     d_anio = Year(KEY_FECHA)
     d_ruc = KEY_RUC
     o_Hoja.Cells(Fila + 1, columna).Value = "FORMATO 13.1: REGISTRO DE INVENTARIO PERMANENTE VALORIZADO-DETALLE DE INVENTARIO VALORIZADO"
  End If
    
    
    o_Hoja.Cells(Fila + 2, columna).Value = "Periodo: " & d_periodo + Space(2) + str(d_anio)
    o_Hoja.Cells(Fila + 3, columna).Value = "RUC: " & KEY_RUC
    o_Hoja.Cells(Fila + 4, columna).Value = "Apellidos y Nombres, Denominaci�n o Raz�n Social: " & KEY_EMPRESA
    
    
    ' -- Obtencion de datos para exportacion
    Dim SqlDatos As String
    Dim rsDatos As ADODB.Recordset
    
    Dim numreg As Long
    Dim TotalRegs As Long
    Dim DivReg As Long
    
    Dim RegsDiv As Boolean
    Dim RegsFinal As Boolean
    If FrmRegistroVentas.Procedencia = nuevo Then
        TipoReg = 1
    End If
    If FrmRegistroCompras.Procedencia = nuevo Then
        TipoReg = 0
    End If
    
    If FrmKardexdeProductos.Procedencia = nuevo Then
        TipoReg = 10
    End If
    
    
    Select Case TipoReg 'Define Flags de columnas
    '01: Ventas
    '02:Compras
    
        Case 1
            SqlDatos = "SELECT * FROM registroventassunat WHERE ruc='" & KEY_RUC & "' AND mes='" & FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1) & "' AND anio='" & FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3) & "'    AND dni_save='" & KEY_USUARIO & "'  ORDER BY fecha,id_doc,serie,numero ASC"
        Case 0
           SqlDatos = "SELECT * FROM movimiento_compra WHERE ruc='" & KEY_RUC & "' AND id_periodo='" & Trim(Me.txtidperiodo.Text) & "' AND id_doc<>'0089'  ORDER BY fecha_emision,id_doc,serie,numero ASC"
        Case 3
            SqlDatos = "SELECT * FROM movimiento_compra C,persona P WHERE C.id_proveedor=P.dni AND C.ruc='" & KEY_RUC & "' AND C.fecha_emision>='" & Format(FrmDaotCompra.DtpInicio.Value, "YYYY-mm-dd") & "' AND C.fecha_emision<='" & Format(FrmDaotCompra.DtpFinal.Value, "YYYY-mm-dd") & "' AND anulado='no' AND id_doc<>'0089'"
        Case 6
            SqlDatos = "CALL sp_temp_reg_grep(" & IdEmpPeriodo & ")"
        Case 10
            SqlDatos = "SELECT * FROM view_kardex WHERE id_alm='" & FrmKardexdeProductos.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(FrmKardexdeProductos.txtcodigoprod.Text) & "' AND ruc='" & KEY_RUC & "'"
    End Select
    
    
    Set rsDatos = ConsultaRegsOpen(SqlDatos)
    
    Fila = Fila + 6
    FactorZoom = 100
    
    If chkNRegs.Value Then
        If Val(txtNRegs.Text) > 0 Then
            DivReg = Val(txtNRegs.Text)
            If DivReg >= 30 And DivReg <= 100 Then
                'Funcion de aproximaci�n al zoom ideal en esa cantidad de lineas por pagina
                FactorZoom = Round(((((DivReg ^ 2 * 10 ^ (8 / DivReg) - DivReg ^ 2 + 10 ^ 3) / (DivReg * 10 ^ 2)) + (DivReg / 10 ^ 6)) * (216 - DivReg)), 0)
            End If
        End If
    End If
    
    If Not rsDatos.EOF Then
        TotalRegs = rsDatos.RecordCount
        pbrAvance.Min = 0
        pbrAvance.Max = TotalRegs
    End If
    
    numreg = 0

    If optTodos.Value Then 'Define Flags de columnas
    
        Select Case TipoReg
            Case 1
                FlagExport = True
                FlagAfecto = True
                FlagExonerado = True
                FlagInafecto = True
                FlagIGV = True
                FlagBIvap = True
                FlagIvap = True
                FlagRetencion = True
                FlagPercepcion = False
                FlagNaturaleza = False
             Case 10
                FlagExport = True
                FlagAfecto = True
                FlagExonerado = True
                FlagInafecto = True
                FlagIGV = True
                FlagBIvap = True
                FlagIvap = True
                FlagRetencion = True
                FlagPercepcion = False
                FlagNaturaleza = False
            Case 0
                FlagAnDua = True
                FlagMGrv1 = True
                FlagMIgv1 = True
                FlagMGrv2 = True
                FlagMIgv2 = True
                FlagMGrv3 = True
                FlagMIgv3 = True
                FlagNoGrav = True
                FlagNoDom = True
                FlagDetracNum = True
                FlagDetracFec = True
                FlagNaturaleza = True
                FlagCcostos1 = True
                FlagCcostos2 = True
                FlagCcostos3 = True
                FlagPercepcion = True
            Case 6
                FlagIGV = True
                FlagDetalle = True
                FlagAfecto = True
                FlagInafecto = True
                FlagCuenta = True
                FlagCCostos = True
          Case 3
                FlagIGV = True
                FlagDetalle = True
                FlagAfecto = True
                FlagInafecto = True
                FlagCuenta = True
                FlagCCostos = True
        End Select
        
        FlagFecVec = True
        
        If TipoReg <> 6 Then
            FlagISC = True
            FlagPerRet = True
        End If
        
        FlagOtros = True
        FlagTotal = True
        FlagTipoCambio = True
        FlagRFecha = True
        FlagRTipo = True
        FlagRSerie = True
        FlagRNumero = True
        

    ElseIf optCustom.Value Then
        Dim i As Long
        
        For i = 0 To lstDestino.ListCount - 1
        
            Select Case lstDestino.List(i)
                Case "Fec.Vec.": FlagFecVec = True
                Case "Exportaci�n": FlagExport = True
                Case "Afecto": FlagAfecto = True
                Case "Exonerado": FlagExonerado = True
                Case "Inafecto": FlagInafecto = True
                Case "ISC": FlagISC = True
                Case "IGV": FlagIGV = True
                Case "Base IVAP": FlagBIvap = True
                Case "IVAP": FlagIvap = True
                Case "Retenci�n": FlagPerRet = True
                Case "A�o DUA": FlagAnDua = True
                Case "M.Gravado1": FlagMGrv1 = True
                Case "IGV 1": FlagMIgv1 = True
                Case "M.Gravado2": FlagMGrv2 = True
                Case "IGV 2": FlagMIgv2 = True
                Case "M.Gravado3": FlagMGrv3 = True
                Case "IGV 3": FlagMIgv3 = True
                Case "Percepci�n": FlagPerRet = True
                Case "N.Comp. No Domiciliado": FlagNoDom = True
                Case "N.Const.Detracci�n": FlagDetracNum = True
                Case "Fec.Const.Detracci�n": FlagDetracFec = True
                
                Case "Detalle": FlagDetalle = True
                Case "Cuenta": FlagCuenta = True
                Case "Centro Costos": FlagCCostos = True
            
                Case "Otros": FlagOtros = True
                Case "Total": FlagTotal = True
                Case "Tipo Cambio": FlagTipoCambio = True
                Case "RD Fecha": FlagRFecha = True
                Case "RD Tipo": FlagRTipo = True
                Case "RD Serie": FlagRSerie = True
                Case "RD Numero": FlagRNumero = True
            End Select
        Next i

        
    ElseIf optSoloValores.Value Then
        Do While Not rsDatos.EOF
            numreg = numreg + 1
            
            Select Case TipoReg
                Case 1
                    
                    If rsDatos!MExport <> 0 Then FlagExport = True
                    If rsDatos!MAfecto <> 0 Then FlagAfecto = True
                    If rsDatos!MExonerado <> 0 Then FlagExonerado = True
                    If rsDatos!MInafecto <> 0 Then FlagInafecto = True
                    If rsDatos!Migv <> 0 Then FlagIGV = True
                    If rsDatos!MBIvap <> 0 Then FlagBIvap = True
                    If rsDatos!MIvap <> 0 Then FlagIvap = True
                    
                Case 0
                    
                    If rsDatos!AnDua <> 0 Then FlagAnDua = True
                    If rsDatos!MGrv1 <> 0 Then FlagMGrv1 = True
                    If rsDatos!MIgv1 <> 0 Then FlagMIgv1 = True
                    If rsDatos!MGrv2 <> 0 Then FlagMGrv2 = True
                    If rsDatos!MIgv2 <> 0 Then FlagMIgv2 = True
                    If rsDatos!MGrv3 <> 0 Then FlagMGrv3 = True
                    If rsDatos!MIgv3 <> 0 Then FlagMIgv3 = True
                    If rsDatos!MNoGrav <> 0 Then FlagNoGrav = True
                    If Not IsNull(rsDatos!NumNoDomiciliado) Then FlagNoDom = True
                    If Not IsNull(rsDatos!DetracNum) Then FlagDetracNum = True
                    If Not IsNull(rsDatos!DetracFec) Then FlagDetracFec = True
                    
                Case 6
                    If Trim(rsDatos!RGDetalle) <> vbNullString Then FlagDetalle = True
                    If rsDatos!MAfecto <> 0 Then FlagAfecto = True
                    If rsDatos!MInafecto <> 0 Then FlagInafecto = True
                    If Not IsNull(rsDatos!NCuenta) Then FlagCuenta = True
                    If rsDatos!Migv <> 0 Then FlagIGV = True
                    
            End Select
        
            If Not IsNull(rsDatos!RVFechaVencimiento) Then FlagFecVec = True
            
            If TipoReg <> 6 Then
                If rsDatos!Misc <> 0 Then FlagISC = True
                If rsDatos!MPerRet <> 0 Then FlagPerRet = True
            End If
            
            If rsDatos!MOtros <> 0 Then FlagOtros = True
            FlagTotal = True
            If rsDatos!TipoCambio <> 0 Then FlagTipoCambio = True
            If Not IsNull(rsDatos!RDFecha) Then FlagRFecha = True
            If Not IsNull(rsDatos!RDTipo) Then FlagRTipo = True
            If Not IsNull(rsDatos!RDSerie) Then FlagRSerie = True
            If Not IsNull(rsDatos!RDNumero) Then FlagRNumero = True
            
            rsDatos.MoveNext
            
            pbrAvance.Value = numreg
            DoEvents
        Loop
        
        rsDatos.MoveFirst
        
    End If
    
    
    Call EscribeCabeceras(o_Hoja, Fila, columna)

    Fila = Fila + 3
    numreg = 0
    
    IniSeccionData = Fila
    Dim Texport As Double
    Dim TBimponible As Double
    Dim Texonerada As Double
    Dim Tinafecta As Double
    Dim tafecto As Double
    Dim Tisc As Double
    Dim tigv As Double
    Dim TbimponibleIvap As Double
    Dim TIvap As Double
    Dim tRetencion As Double
    Dim tpercepcion As Double
    Dim Totros As Double
    Dim tTotal As Double
        
    
    Do While Not rsDatos.EOF
        numreg = numreg + 1
        
        RegsDiv = False
        RegsFinal = False
        
        If numreg = TotalRegs Then
            RegsFinal = True
        Else
            If DivReg > 0 Then
                If ((numreg + 9) Mod DivReg) = 0 Then
                    RegsDiv = True
                End If
            End If
        End If
        
        columna = 1
        o_Hoja.Cells(Fila, columna).Value = numreg
        
        columna = columna + 1
        
        o_Hoja.Cells(Fila, columna).NumberFormat = "dd/mm/yyyy"
        If TipoReg = 1 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha
        Else
            o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha_emision
        End If
        columna = columna + 1

        If FlagFecVec Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "dd/mm/yyyy"
            'o_Hoja.Cells(Fila, Columna).Value = rsDatos!RVFechaVencimiento
            If TipoReg = 1 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha
            End If
            If TipoReg = 0 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha_cancelacion
            End If
            columna = columna + 1
        End If
        
        o_Hoja.Cells(Fila, columna).NumberFormat = "00"
        o_Hoja.Cells(Fila, columna).Value = rsDatos!id_doc
        columna = columna + 1
        
        o_Hoja.Cells(Fila, columna).NumberFormat = "0000"
        o_Hoja.Cells(Fila, columna).Value = rsDatos!serie
        columna = columna + 1
        
        If FlagAnDua Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!anio_dua
            columna = columna + 1
        End If
        
        If TipoReg <> 1 Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "000000"
        End If
        o_Hoja.Cells(Fila, columna).Value = rsDatos!numero
        columna = columna + 1
        If TipoReg = 0 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!tipo_doc_identidad
        Else
            o_Hoja.Cells(Fila, columna).Value = rsDatos!tdoc
        End If
        columna = columna + 1
        If TipoReg = 0 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!id_proveedor
        Else
            o_Hoja.Cells(Fila, columna).Value = rsDatos!ruccliente
        End If
        columna = columna + 1
         If TipoReg = 0 Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!nproveedor
        Else
            o_Hoja.Cells(Fila, columna).Value = rsDatos!NombreCliente
        End If
        
        
        If RegsFinal Then
            o_Hoja.Cells(Fila + 1, columna).Value = "TOTALES"
            o_Hoja.Cells(Fila + 1, columna).Font.Bold = True
        ElseIf RegsDiv Then
            o_Hoja.Cells(Fila + 1, columna).Value = "VAN"
            o_Hoja.Cells(Fila + 1, columna).Font.Bold = True
             Fila = Fila + 5
            Call EscribeCabeceras(o_Hoja, Fila, 1)

           
            o_Hoja.Cells(Fila + 3, columna).Value = "VIENEN"
            o_Hoja.Cells(Fila + 3, columna).Font.Bold = True
        End If
        
        columna = columna + 1
        
        '==== Secci�n exclusiva gastos =====
        If FlagDetalle Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RGDetalle
            columna = columna + 1
        End If
    
        If FlagInafecto And TipoReg = 6 Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!MInafecto
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = rsDatos!Tinafecto
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 4, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 4, columna).Value = rsDatos!Tinafecto
                End If
            End If
            columna = columna + 1
        End If
        
        '==== Secci�n exclusiva ventas =====
        If FlagExport Then
            If RegsDiv Then
           Fila = Fila - 5
           End If
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!MExport
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Texport 'rsDatos!TExport
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Texport ' rsDatos!TExport
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagAfecto Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!afecto) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!afecto
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            tafecto = tafecto + rsDatos!afecto
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = tafecto 'rsDatos!afecto
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = tafecto 'rsDatos!afecto
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagExonerado Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!exonerado) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!exonerado
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            Texonerada = Texonerada + rsDatos!exonerado
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Texonerada 'rsDatos!exonerado
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Texonerada 'rsDatos!exonerado
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagInafecto And TipoReg <> 6 Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!exonerado
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Tinafecta ' rsDatos!TInafecto
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Tinafecta 'rsDatos!TInafecto
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagISC And TipoReg = 1 Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!Misc
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Tisc ' rsDatos!TIsc
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Tisc 'rsDatos!TIsc
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagIGV Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!igv) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!igv
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            
            tigv = tigv + rsDatos!igv
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = tigv 'rsDatos!tigv
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = tigv 'rsDatos!tigv
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagBIvap Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!MBIvap
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TIvap 'rsDatos!TBIvap
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TIvap ' rsDatos!TBIvap
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagIvap Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = "" ' rsDatos!MIvap
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TIvap ' rsDatos!TIvap
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TIvap 'rsDatos!TIvap
                End If
            End If
            columna = columna + 1
        End If


        '==== Secci�n exclusiva compras =====
        If TipoReg = 0 Then
             If RegsDiv Then
                Fila = Fila - 5
             End If
        If (FlagMGrv1 Or FlagMIgv1) And rsDatos!id_tipo_compra = "01" Then
           
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If rsDatos!valor_venta <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!valor_venta
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTGrv1 = TTGrv1 + rsDatos!valor_venta
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TTGrv1 'rsDatos!TGrv1
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TTGrv1 'rsDatos!tGrv1
                End If
            End If
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!igv) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!igv
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTigv1 = TTigv1 + rsDatos!igv
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TTigv1 'rsDatos!tigv1
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TTigv1 'rsDatos!tigv1
                End If
            columna = columna + 1
            Else
            columna = columna + 5
            End If
            
        Else
        If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTGrv1) 'rsDatos!tigv1
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTGrv1) 'rsDatos!tigv1
                End If
                columna = columna + 1
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTigv1)
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTigv1)
                End If
                 columna = columna + 1
       End If
        End If
    
        
        
        
        
        
        If (FlagMGrv2 Or FlagMIgv2) And rsDatos!id_tipo_compra = "02" Then
              If RegsFinal Or RegsDiv Then
                columna = columna
              Else
                columna = columna + 2
              End If
             
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!valor_venta) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!valor_venta
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTGrv2 = TTGrv2 + rsDatos!valor_venta
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTGrv2) 'rsDatos!TGrv2
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTGrv2) 'rsDatos!TGrv2
                End If
            End If
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!igv) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!igv
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTigv2 = TTigv2 + rsDatos!igv
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTigv2) 'rsDatos!tigv2
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TTigv2 'rsDatos!tigv2
                End If
            columna = columna + 1
            Else
            columna = columna + 3
            End If
            
        Else
        If RegsFinal Or RegsDiv Then
                
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTGrv2)
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTGrv2)
                End If
                columna = columna + 1
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTigv2)
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTigv2)
                End If
                columna = columna + 1
        End If
       End If
        
        
        
        If (FlagMGrv3 Or FlagMIgv3) And rsDatos!id_tipo_compra = "03" Then
           If RegsFinal Or RegsDiv Then
                columna = columna
           Else
                columna = columna + 4
           End If
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!valor_venta) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!valor_venta
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTGrav3 = TTGrav3 + rsDatos!valor_venta
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTGrav3) 'rsDatos!TGrv3
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTGrav3) 'rsDatos!TGrv3
                End If
            End If
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!igv) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!igv
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTigv3 = TTigv3 + rsDatos!igv
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTigv3) 'rsDatos!tigv3
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTigv3) 'rsDatos!tigv3
                End If
                columna = columna + 1
            Else
                columna = columna + 1
                
            End If
            
         Else
         If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTGrav3)
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTGrav3)
                End If
                columna = columna + 1
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(TTigv3)
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(TTigv3)
                End If
                 columna = columna + 1
        End If
        End If
        
        
        
        
        
        
        
        
        
        If FlagNoGrav Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!exonerado) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!exonerado
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTnograv = TTnograv + rsDatos!exonerado
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TTnograv 'rsDatos!TNoGrav
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TTnograv 'rsDatos!TNoGrav
                End If
            End If
            columna = columna + 1
        End If
    
        If FlagISC And TipoReg <> 1 Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!isc) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!isc
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            TTisc = TTisc + rsDatos!isc
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = TTisc 'rsDatos!Tisc
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = TTisc 'rsDatos!Tisc
                End If
            End If
            columna = columna + 1
        End If
    End If
        '====================================
    
        If FlagPercepcion Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!percepcion) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!percepcion
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            tpercepcion = tpercepcion + rsDatos!percepcion
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = tpercepcion 'rsDatos!TPerRet
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = tpercepcion 'rsDatos!TPerRet
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagRetencion Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!retencion) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!retencion
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            tRetencion = tRetencion + rsDatos!retencion
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = tRetencion 'rsDatos!TPerRet
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = tRetencion 'rsDatos!TPerRet
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagOtros Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            If Val(rsDatos!Otros) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!Otros
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = Val(Totros) 'rsDatos!TOtros
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = Val(Totros) ' rsDatos!TOtros
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagTotal Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "#,##0.00"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!Total
            tTotal = tTotal + rsDatos!Total
            If RegsFinal Or RegsDiv Then
                o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                o_Hoja.Cells(Fila + 1, columna).Value = tTotal 'rsDatos!tTotal
                If RegsDiv Then
                    o_Hoja.Cells(Fila + 8, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 8, columna).Value = tTotal ' rsDatos!tTotal
                End If
            End If
            columna = columna + 1
        End If
        
        If FlagNoDom Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "00000000"
            If rsDatos!numero_no_domiciliado <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!numero_no_domiciliado
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
        End If
        
        If FlagDetracNum Or FlagDetracFec Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "00000000"
            If rsDatos!numero_detrac <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!numero_detrac
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "dd/mm/yyyy"
            If Val(rsDatos!numero_detrac) <> 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha_detrac
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
        End If
    
        If FlagTipoCambio Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "0.000"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!tc
            columna = columna + 1
        End If
        
       
        
        If FlagRFecha Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "dd/mm/yyyy"
            If TipoReg = 1 Then
            If Val(rsDatos!numeroF) > 0 And Val(rsDatos!doc_codF) > 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha_F
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            
            columna = columna + 1
        End If
        
        If FlagRTipo Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "00"
            If TipoReg = 1 Then
            If Val(rsDatos!doc_codF) > 0 Then
                o_Hoja.Cells(Fila, columna).Value = Mid(rsDatos!doc_codF, 2, 3)
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            Else
            o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
        End If
        
        If FlagRSerie Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "0000"
             If TipoReg = 1 Then
            If Val(rsDatos!doc_codF) > 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!serieF
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            Else
            o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
        End If
        
        If FlagRNumero Then
            o_Hoja.Cells(Fila, columna).NumberFormat = "00000000"
            If TipoReg = 1 Then
            If Val(rsDatos!doc_codF) > 0 Then
                o_Hoja.Cells(Fila, columna).Value = rsDatos!numeroF
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
        End If
        
        If FlagCuenta Then
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!NCuenta
            columna = columna + 1
        End If
        If FlagNaturaleza Then
            strCadena = "SELECT * FROM movimiento_compra_costos_naturaleza WHERE id_compra='" & rsDatos("id_compra") & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                o_Hoja.Cells(Fila, columna).Value = rstT("cnaturaleza")
            Else
                o_Hoja.Cells(Fila, columna).Value = ""
            End If
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
            o_Hoja.Cells(Fila, columna).Value = "" 'rsDatos!total - rsDatos!igv
            columna = columna + 1
            strCadena = "SELECT * FROM movimiento_compra_naturaleza_costos WHERE id_compra='" & Val(rsDatos("id_compra")) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
            rstT.MoveFirst
            For q = 0 To rstT.RecordCount - 1
                o_Hoja.Cells(Fila, columna).Value = rstT("ccostos")
                columna = columna + 1
                  o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rstT("monto")
                columna = columna + 1
                rstT.MoveNext
            Next q
            End If
        End If
        If FlagCCostos Then
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCGastosAdmin
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCVentas
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCProduccion
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCExcepcionales
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCDiversos
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!RVCCFinancieros
            columna = columna + 1
        End If
    
        
        rsDatos.MoveNext
        pbrAvance.Value = numreg
        DoEvents
        
        If Not RegsFinal Then
            If RegsDiv Then
                Call SetEstiloGenData(o_Hoja, IniSeccionData, 1, Fila, UltColumna)
                Call SetEstiloGenTotales(o_Hoja, Fila + 1, ColNomEntidad, Fila + 1, UltColumna)
                Call SetEstiloGenTotales(o_Hoja, Fila + 8, ColNomEntidad, Fila + 8, UltColumna)
                Fila = Fila + 9
                IniSeccionData = Fila
            Else
                Fila = Fila + 1
            End If
        Else
            Call SetEstiloGenData(o_Hoja, IniSeccionData, 1, Fila, UltColumna)
            Call SetEstiloGenTotales(o_Hoja, Fila + 1, ColNomEntidad, Fila + 1, UltColumna)
            o_Hoja.Range(o_Hoja.Cells(FinFilaCab, 1), o_Hoja.Cells(Fila, UltColumna)).AutoFilter
            If TipoReg = 1 Then
                o_Hoja.Columns(ColNumero).EntireColumn.AutoFit
            End If
        End If
        
    Loop
    
    Set rsDatos = Nothing
    
    ' -- Atributos visuales de la hoja
    With o_Excel.ActiveWindow
        .DisplayGridlines = False
        .Zoom = 80
    End With
    
    ' -- Setup de impresi�n de la hoja
    With o_Hoja.PageSetup
        .CenterFooter = PiePagina
        .LeftMargin = o_Excel.InchesToPoints(0.984251968503937)
        .RightMargin = o_Excel.InchesToPoints(0.196850393700787)
        .TopMargin = o_Excel.InchesToPoints(0.47244094488189)
        .BottomMargin = o_Excel.InchesToPoints(0.31496062992126)
        .HeaderMargin = o_Excel.InchesToPoints(0.47244094488189)
        .FooterMargin = o_Excel.InchesToPoints(0.15748031496063)
        .PaperSize = xlPaperLegal
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .Zoom = FactorZoom
    End With
    
    
    o_Libro.close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
'error_Handler:
    ' -- Cierra la hoja y el la aplicaci�n Excel
 '   If Not o_Libro Is Nothing Then: o_Libro.Close False
  '  If Not o_Excel Is Nothing Then: o_Excel.Quit
   ' Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
   ' MsgBox Err.Description, vbCritical
End Function

Private Function Exportar_Excel_vv(sOutputPath As String) As Boolean
  
   ' On Error GoTo error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim iSheetsPerBook As Integer
    
    Dim Fila        As Long
    Dim columna     As Long
    
    Dim IniSeccionData As Long
    
    Dim FactorZoom As Integer
    
    Dim in_producto As String
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    
    iSheetsPerBook = o_Excel.SheetsInNewWorkbook
    o_Excel.SheetsInNewWorkbook = 1
    
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets(1)
    o_Hoja.name = "Reporte"
    
    o_Excel.SheetsInNewWorkbook = iSheetsPerBook
    
    Fila = 1
    columna = 1
    
    ' -- Grabar Datos de Cabecera
    o_Hoja.Cells(Fila, columna).Value = FormatNom
    o_Hoja.Cells(Fila, columna).Font.Size = 13
    o_Hoja.Cells(Fila, columna).Font.Bold = True
    
   If FrmRegistroVentas.Procedencia = nuevo Then
     If Me.chkSucursal.Value = 1 Then
        in_sucursal = Me.DtcSucursal.Text
    Else
        in_sucursal = ""
     End If
     d_periodo = "PERIODO" & Mid(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 2), 18, 15) + Space(1) + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3) & Space(1) + in_sucursal
     d_anio = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)
     d_ruc = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 0)
     o_Hoja.Cells(Fila + 1, columna).Value = "FORMATO 14.1: REGISTRO DE VENTAS E INGRESOS"
     o_Hoja.Cells(Fila + 2, columna).Value = "Periodo: " & d_periodo + Space(2) + str(d_anio)
     o_Hoja.Cells(Fila + 3, columna).Value = "RUC: " & KEY_RUC
     o_Hoja.Cells(Fila + 4, columna).Value = "Apellidos y Nombres, Denominaci�n o Raz�n Social: " & KEY_EMPRESA
  End If
     
  If FrmRegistroCompras.Procedencia = nuevo Then
     d_periodo = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2)
     d_anio = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
     d_ruc = FrmRegistroCompras.txtRuc.Text
     o_Hoja.Cells(Fila + 1, columna).Value = "FORMATO 8.4: REGISTRO DE COMPRAS"
     
     o_Hoja.Cells(Fila + 2, columna).Value = "Periodo: " & d_periodo + Space(2) + str(d_anio)
    o_Hoja.Cells(Fila + 3, columna).Value = "RUC: " & KEY_RUC
    o_Hoja.Cells(Fila + 4, columna).Value = "Apellidos y Nombres, Denominaci�n o Raz�n Social: " & KEY_EMPRESA
   End If
  
  If FrmDaotCompra.Procedencia = nuevo Then
     d_periodo = str(FrmDaotCompra.DtpInicio.Value) + "-" + str(FrmDaotCompra.DtpFinal.Value)
     d_anio = Year(KEY_FECHA)
     d_ruc = KEY_RUC
     
     o_Hoja.Cells(Fila + 2, columna).Value = "Periodo: " & d_periodo + Space(2) + str(d_anio)
    o_Hoja.Cells(Fila + 3, columna).Value = "RUC: " & KEY_RUC
    o_Hoja.Cells(Fila + 4, columna).Value = "Apellidos y Nombres, Denominaci�n o Raz�n Social: " & KEY_EMPRESA
    
  End If
  
  If FrmKardexdeProductos.Procedencia = nuevo Then
     If Me.chkSucursal.Value = 1 Then
        in_sucursal = Me.DtcSucursal.Text
     Else
        in_sucursal = ""
     End If
     d_periodo = "PERIODO:"
     d_anio = Year(KEY_FECHA)
     d_ruc = KEY_RUC
    
    o_Hoja.Cells(Fila + 1, columna).Value = "RUC: "
    o_Hoja.Cells(Fila + 1, columna + 1).Value = KEY_RUC
    o_Hoja.Cells(Fila + 2, columna).Value = "RAZON SOCIAL:"
    o_Hoja.Cells(Fila + 2, columna + 1).Value = KEY_EMPRESA
    
    o_Hoja.Cells(Fila + 3, columna).Value = "FORMATO 13.1:"
    o_Hoja.Cells(Fila + 3, columna + 1).Value = "REGISTRO DE INVENTARIO PERMANENTE VALORIZADO-DETALLE DE INVENTARIO VALORIZADO"
    o_Hoja.Cells(Fila + 4, columna).Value = "PERIODO: "
    o_Hoja.Cells(Fila + 4, columna + 1).Value = Format(FrmKardexdeProductos.DtpDesde.Value, "dd-mm-YYYY") + Space(2) + Format(FrmKardexdeProductos.DtpHasta.Value, "dd-mm-YYYY")
    o_Hoja.Cells(Fila + 5, columna).Value = "METODO DE VALUACION  :"
    o_Hoja.Cells(Fila + 5, columna + 1).Value = "COSTO PROMEDIO"
    o_Hoja.Cells(Fila + 6, columna).Value = "EXPRESADO EN SOLES"
    
    If FrmKardexdeProductos.chk_all.Value = 1 Then
       o_Hoja.Cells(Fila + 7, columna).Value = "[TODOS LOS ALMACENES]"
    Else
       o_Hoja.Cells(Fila + 7, columna).Value = FrmKardexdeProductos.DtcAlmacen.Text
    End If
    
    
  End If
    
    
    
    
    
    ' -- Obtencion de datos para exportacion
    Dim SqlDatos As String
    Dim rsDatos As ADODB.Recordset
    
    Dim numreg As Long
    Dim TotalRegs As Long
    Dim DivReg As Long
    
    Dim RegsDiv As Boolean
    Dim RegsFinal As Boolean
    If FrmRegistroVentas.Procedencia = nuevo Then
        TipoReg = 1
    End If
    If FrmRegistroCompras.Procedencia = nuevo Then
        TipoReg = 0
    End If
    
    If FrmKardexdeProductos.Procedencia = nuevo Then
        TipoReg = 10
    End If
    
    Select Case TipoReg 'Define Flags de columnas
    '1: Ventas
    '0:Compras
        Case 1
            SqlDatos = "SELECT * FROM registroventassunat WHERE ruc='" & KEY_RUC & "' AND mes='" & FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1) & "' AND anio='" & FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3) & "'    AND dni_save='" & KEY_USUARIO & "'  ORDER BY fecha,id_doc,serie,numero ASC"
        Case 0
            SqlDatos = "SELECT * FROM movimiento_compra WHERE ruc='" & KEY_RUC & "' AND id_periodo='" & Trim(Me.txtidperiodo.Text) & "' AND id_doc not in('0089','0090')  ORDER BY fecha_emision,id_doc,serie,numero ASC"
        Case 3
            SqlDatos = "SELECT * FROM movimiento_compra C,persona P WHERE C.id_proveedor=P.dni AND C.ruc='" & KEY_RUC & "' AND C.fecha_emision>='" & Format(FrmDaotCompra.DtpInicio.Value, "YYYY-mm-dd") & "' AND C.fecha_emision<='" & Format(FrmDaotCompra.DtpFinal.Value, "YYYY-mm-dd") & "' AND anulado='no' AND id_doc not in('0089','0090')"
        Case 6
            SqlDatos = "CALL sp_temp_reg_grep(" & IdEmpPeriodo & ")"
        Case 10
            If FrmKardexdeProductos.txtcodigoprod.Text <> "" Then
               If FrmKardexdeProductos.chkBuscarfechas.Value = 1 Then
                SqlDatos = "SELECT * FROM view_kardex_sunat_v2 WHERE  fecha_emision>='" & Format(FrmKardexdeProductos.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(FrmKardexdeProductos.DtpHasta.Value, "YYYY-mm-dd") & "' and id_alm='" & FrmKardexdeProductos.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(FrmKardexdeProductos.txtcodigoprod.Text) & "' AND ruc='" & KEY_RUC & "'"
            Else
                SqlDatos = "SELECT * FROM view_kardex_sunat_v2 WHERE   id_alm='" & FrmKardexdeProductos.DtcAlmacen.BoundText & "' AND id_producto='" & Trim(FrmKardexdeProductos.txtcodigoprod.Text) & "' AND ruc='" & KEY_RUC & "'"
            End If
            Else
               SqlDatos = "SELECT * FROM view_kardex_sunat_v2 WHERE  fecha_emision>='" & Format(FrmKardexdeProductos.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(FrmKardexdeProductos.DtpHasta.Value, "YYYY-mm-dd") & "' and id_alm='" & FrmKardexdeProductos.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "'"
            End If
            
    End Select
    
    Set rsDatos = ConsultaRegsOpen(SqlDatos)
    
   Fila = Fila + 6
    
    FactorZoom = 100
    
    If chkNRegs.Value Then
        If Val(txtNRegs.Text) > 0 Then
            DivReg = Val(txtNRegs.Text)
            If DivReg >= 30 And DivReg <= 100 Then
                'Funcion de aproximaci�n al zoom ideal en esa cantidad de lineas por pagina
                FactorZoom = Round(((((DivReg ^ 2 * 10 ^ (8 / DivReg) - DivReg ^ 2 + 10 ^ 3) / (DivReg * 10 ^ 2)) + (DivReg / 10 ^ 6)) * (216 - DivReg)), 0)
            End If
        End If
    End If
    
    If Not rsDatos.EOF Then
        TotalRegs = rsDatos.RecordCount
        pbrAvance.Min = 0
        pbrAvance.Max = TotalRegs
    End If
    
    numreg = 0

    If optTodos.Value Then 'Define Flags de columnas
    
        Select Case TipoReg
            Case 1
                FlagExport = True
                FlagAfecto = True
                FlagExonerado = True
                FlagInafecto = True
                FlagIGV = True
                FlagBIvap = True
                FlagIvap = True
                FlagRetencion = True
                FlagPercepcion = False
                FlagNaturaleza = False
             Case 10
                FlagAfecto = False
                FlagFecVec = False
                FlagTipoOpe = True
                FlagCantIng = True
                FlagCostoIng = True
                FlagSaldoIng = True
                FlagCantSal = True
                FlagCostoSal = True
                FlagSaldoSalida = True
                FlagCantSaldo = True
                FlagCostoPomedio = True
                FlagSaldoFinal = True
                FlagPercepcion = False
                FlagNaturaleza = False
                FlagAfecto = False
                FlagExonerado = False
                FlagInafecto = False
                FlagIGV = False
                FlagBIvap = False
                FlagIvap = False
                FlagFecVec = False
                
            Case 0
                FlagAnDua = True
                FlagMGrv1 = True
                FlagMIgv1 = True
                FlagMGrv2 = True
                FlagMIgv2 = True
                FlagMGrv3 = True
                FlagMIgv3 = True
                FlagNoGrav = True
                FlagNoDom = True
                FlagDetracNum = True
                FlagDetracFec = True
                FlagNaturaleza = True
                FlagCcostos1 = True
                FlagCcostos2 = True
                FlagCcostos3 = True
                FlagPercepcion = True
            Case 6
                FlagIGV = True
                FlagDetalle = True
                FlagAfecto = True
                FlagInafecto = True
                FlagCuenta = True
                FlagCCostos = True
          Case 3
                FlagIGV = True
                FlagDetalle = True
                FlagAfecto = True
                FlagInafecto = True
                FlagCuenta = True
                FlagCCostos = True
        End Select
        
        FlagFecVec = True
        
        If TipoReg = 0 Or TipoReg = 1 Then
            FlagISC = True
            FlagPerRet = True
        End If
        
        
        
        
        
        FlagOtros = True
        FlagTotal = True
        FlagTipoCambio = True
        FlagRFecha = True
        FlagRTipo = True
        FlagRSerie = True
        FlagRNumero = True
        

    ElseIf optCustom.Value Then
        Dim i As Long
        
        For i = 0 To lstDestino.ListCount - 1
        
            Select Case lstDestino.List(i)
                Case "Fec.Vec.": FlagFecVec = True
                
                Case "Exportaci�n": FlagExport = True
                Case "Afecto": FlagAfecto = True
                Case "Exonerado": FlagExonerado = True
                Case "Inafecto": FlagInafecto = True
                Case "ISC": FlagISC = True
                Case "IGV": FlagIGV = True
                Case "Base IVAP": FlagBIvap = True
                Case "IVAP": FlagIvap = True
                Case "Retenci�n": FlagPerRet = True
                
                Case "A�o DUA": FlagAnDua = True
                Case "M.Gravado1": FlagMGrv1 = True
                Case "IGV 1": FlagMIgv1 = True
                Case "M.Gravado2": FlagMGrv2 = True
                Case "IGV 2": FlagMIgv2 = True
                Case "M.Gravado3": FlagMGrv3 = True
                Case "IGV 3": FlagMIgv3 = True
                Case "Percepci�n": FlagPerRet = True
                Case "N.Comp. No Domiciliado": FlagNoDom = True
                Case "N.Const.Detracci�n": FlagDetracNum = True
                Case "Fec.Const.Detracci�n": FlagDetracFec = True
                
                Case "Detalle": FlagDetalle = True
                Case "Cuenta": FlagCuenta = True
                Case "Centro Costos": FlagCCostos = True
            
                Case "Otros": FlagOtros = True
                Case "Total": FlagTotal = True
                Case "Tipo Cambio": FlagTipoCambio = True
                Case "RD Fecha": FlagRFecha = True
                Case "RD Tipo": FlagRTipo = True
                Case "RD Serie": FlagRSerie = True
                Case "RD Numero": FlagRNumero = True
            End Select
        Next i

        
    ElseIf optSoloValores.Value Then
        Do While Not rsDatos.EOF
            numreg = numreg + 1
            
            Select Case TipoReg
                Case 1
                    If rsDatos!MExport <> 0 Then FlagExport = True
                    If rsDatos!MAfecto <> 0 Then FlagAfecto = True
                    If rsDatos!MExonerado <> 0 Then FlagExonerado = True
                    If rsDatos!MInafecto <> 0 Then FlagInafecto = True
                    If rsDatos!Migv <> 0 Then FlagIGV = True
                    If rsDatos!MBIvap <> 0 Then FlagBIvap = True
                    If rsDatos!MIvap <> 0 Then FlagIvap = True
                    
                Case 0
                    If rsDatos!AnDua <> 0 Then FlagAnDua = True
                    If rsDatos!MGrv1 <> 0 Then FlagMGrv1 = True
                    If rsDatos!MIgv1 <> 0 Then FlagMIgv1 = True
                    If rsDatos!MGrv2 <> 0 Then FlagMGrv2 = True
                    If rsDatos!MIgv2 <> 0 Then FlagMIgv2 = True
                    If rsDatos!MGrv3 <> 0 Then FlagMGrv3 = True
                    If rsDatos!MIgv3 <> 0 Then FlagMIgv3 = True
                    If rsDatos!MNoGrav <> 0 Then FlagNoGrav = True
                    If Not IsNull(rsDatos!NumNoDomiciliado) Then FlagNoDom = True
                    If Not IsNull(rsDatos!DetracNum) Then FlagDetracNum = True
                    If Not IsNull(rsDatos!DetracFec) Then FlagDetracFec = True
                Case 10
                   
                    If rsDatos!MGrv1 <> 0 Then FlagMGrv1 = True
                    If rsDatos!MIgv1 <> 0 Then FlagMIgv1 = True
                    If rsDatos!MGrv2 <> 0 Then FlagMGrv2 = True
                    If rsDatos!MIgv2 <> 0 Then FlagMIgv2 = True
                    If rsDatos!MGrv3 <> 0 Then FlagMGrv3 = True
                    If rsDatos!MIgv3 <> 0 Then FlagMIgv3 = True
                    If rsDatos!MNoGrav <> 0 Then FlagNoGrav = True
                    If Not IsNull(rsDatos!NumNoDomiciliado) Then FlagNoDom = True
                    If Not IsNull(rsDatos!DetracNum) Then FlagDetracNum = True
                    If Not IsNull(rsDatos!DetracFec) Then FlagDetracFec = True
                    
                Case 6
                    If Trim(rsDatos!RGDetalle) <> vbNullString Then FlagDetalle = True
                    If rsDatos!MAfecto <> 0 Then FlagAfecto = True
                    If rsDatos!MInafecto <> 0 Then FlagInafecto = True
                    If Not IsNull(rsDatos!NCuenta) Then FlagCuenta = True
                    If rsDatos!Migv <> 0 Then FlagIGV = True
                    
            End Select
        
            If Not IsNull(rsDatos!RVFechaVencimiento) Then FlagFecVec = True
            
            If TipoReg <> 6 Then
                If rsDatos!Misc <> 0 Then FlagISC = True
                If rsDatos!MPerRet <> 0 Then FlagPerRet = True
            End If
            
            If rsDatos!MOtros <> 0 Then FlagOtros = True
            FlagTotal = True
            If rsDatos!TipoCambio <> 0 Then FlagTipoCambio = True
            If Not IsNull(rsDatos!RDFecha) Then FlagRFecha = True
            If Not IsNull(rsDatos!RDTipo) Then FlagRTipo = True
            If Not IsNull(rsDatos!RDSerie) Then FlagRSerie = True
            If Not IsNull(rsDatos!RDNumero) Then FlagRNumero = True
            
            rsDatos.MoveNext
            
            pbrAvance.Value = numreg
            DoEvents
        Loop
        
        rsDatos.MoveFirst
        in_producto = rsDatos("id_producto")
    End If
    
    If rsDatos.RecordCount > 0 Then
         rsDatos.MoveFirst
        in_producto = rsDatos("id_producto")
    End If
    
    If TipoReg = "10" Then 'KARDEX
        Call EscribeCabecerasKardex(o_Hoja, Fila, columna, rsDatos("id_producto"))
    Else
        Call EscribeCabeceras(o_Hoja, Fila, columna)
    End If
    

   
    numreg = 0

    IniSeccionData = Fila
    Dim Texport As Double
    Dim TBimponible As Double
    Dim Texonerada As Double
    Dim Tinafecta As Double
    Dim tafecto As Double
    Dim Tisc As Double
    Dim tigv As Double
    Dim TbimponibleIvap As Double
    Dim TIvap As Double
    Dim tRetencion As Double
    Dim tpercepcion As Double
    Dim Totros As Double
    Dim tTotal As Double
    Dim T_cant_ingreso As Double
        
 
   
    Do While Not rsDatos.EOF
        numreg = numreg + 1
        
        RegsDiv = False
        RegsFinal = False
        
        If numreg = TotalRegs Then
            RegsFinal = True
        Else
            If DivReg > 0 Then
                If ((numreg + 9) Mod DivReg) = 0 Then
                    RegsDiv = True
                End If
            End If
        End If
        
        columna = 1
 
        If TipoReg = 10 Then
            
            
        If in_producto <> rsDatos("id_producto") Then
            in_producto = rsDatos("id_producto")
           
            Call EscribeCabecerasKardex(o_Hoja, Fila, 1, rsDatos("id_producto"))
            
        End If
 
  
        
            
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "dd/mm/YYYY"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!fecha_emision
            
           
             
        '--
        
            
            
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "00"
            o_Hoja.Cells(Fila, columna).Value = Format(rsDatos!tdoc, "00")
            columna = columna + 1
            
             
            o_Hoja.Cells(Fila, columna).Value = rsDatos!serie
            columna = columna + 1
            o_Hoja.Cells(Fila, columna).Value = rsDatos!numero
             columna = columna + 1
             
            
            
            If rsDatos!cantidad_real > 0 Then
                o_Hoja.Cells(Fila, columna).Value = "COMPRAS"
                T_cant_ingreso = T_cant_ingreso + rsDatos!cantidad_real
                
                If RegsFinal Then
                    o_Hoja.Cells(Fila + 1, columna).Value = "TOTALES"
                    o_Hoja.Cells(Fila + 1, columna).Font.Bold = True
               
                End If
                
                
                
                
                
                columna = 6
                
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!cantidad
                columna = columna + 1
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_unitario
                columna = columna + 1
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_unitario * rsDatos!cantidad
                columna = columna + 4
            
            Else
                o_Hoja.Cells(Fila, columna).Value = "VENTAS"
                 columna = 9
                
                If RegsFinal Then
                    o_Hoja.Cells(Fila + 1, columna).Value = "TOTALES"
                    o_Hoja.Cells(Fila + 1, columna).Font.Bold = True
                
                End If
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!cantidad
                
                If RegsFinal Or RegsDiv Then
                    o_Hoja.Cells(Fila + 1, columna).NumberFormat = "#,##0.00"
                    o_Hoja.Cells(Fila + 1, columna).Value = T_cant_ingreso
                    If RegsDiv Then
                        o_Hoja.Cells(Fila + 4, columna).NumberFormat = "#,##0.00"
                        o_Hoja.Cells(Fila + 4, columna).Value = T_cant_ingreso
                    End If
                End If
            
                
                
                
                columna = columna + 1
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_unitario
                columna = columna + 1
                o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
                o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_unitario * rsDatos!cantidad
                columna = columna + 1
            
            End If
        
        
        
        
            o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!saldo_stock
            columna = columna + 1
            
            o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_promedio
            columna = columna + 1
                
            o_Hoja.Cells(Fila, columna).NumberFormat = "0.00"
            o_Hoja.Cells(Fila, columna).Value = rsDatos!costo_promedio * rsDatos!saldo_stock
            columna = columna + 1
            ColNomEntidad = columna
            
            
        End If
        
        
        
        rsDatos.MoveNext
        pbrAvance.Value = numreg
        DoEvents
        
        If Not RegsFinal Then
            If RegsDiv Then
                Call SetEstiloGenData(o_Hoja, IniSeccionData, 1, Fila, UltColumna)
                Call SetEstiloGenTotales(o_Hoja, Fila + 1, ColNomEntidad, Fila + 1, UltColumna)
                Call SetEstiloGenTotales(o_Hoja, Fila + 8, ColNomEntidad, Fila + 8, UltColumna)
                Fila = Fila + 9
                IniSeccionData = Fila
            Else
                Fila = Fila + 1
            End If
        Else
          '  Call SetEstiloGenData(o_Hoja, IniSeccionData, 1, Fila, UltColumna)
            Call SetEstiloGenTotales(o_Hoja, Fila + 1, ColNomEntidad, Fila + 1, UltColumna)
            o_Hoja.Range(o_Hoja.Cells(FinFilaCab, 1), o_Hoja.Cells(Fila, UltColumna)).AutoFilter
            If TipoReg = 1 Then
                o_Hoja.Columns(ColNumero).EntireColumn.AutoFit
            End If
        End If
        
    Loop
visualizar:
    Set rsDatos = Nothing
    
    ' -- Atributos visuales de la hoja
    With o_Excel.ActiveWindow
        .DisplayGridlines = False
        .Zoom = 80
    End With
    
    ' -- Setup de impresi�n de la hoja
    With o_Hoja.PageSetup
        .CenterFooter = PiePagina
        .LeftMargin = o_Excel.InchesToPoints(0.984251968503937)
        .RightMargin = o_Excel.InchesToPoints(0.196850393700787)
        .TopMargin = o_Excel.InchesToPoints(0.47244094488189)
        .BottomMargin = o_Excel.InchesToPoints(0.31496062992126)
        .HeaderMargin = o_Excel.InchesToPoints(0.47244094488189)
        .FooterMargin = o_Excel.InchesToPoints(0.15748031496063)
        .PaperSize = xlPaperLegal
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .Zoom = FactorZoom
    End With
    
    
    o_Libro.close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel_vv = True
Exit Function
  
' -- Controlador de Errores
'error_Handler:
    ' -- Cierra la hoja y el la aplicaci�n Excel
 '   If Not o_Libro Is Nothing Then: o_Libro.Close False
  '  If Not o_Excel Is Nothing Then: o_Excel.Quit
   ' Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
   ' MsgBox Err.Description, vbCritical
End Function

' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub


Private Sub EscribeCabeceras(ByRef o_Hoja As Object, Fila As Long, columna As Long)

    Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "N�mero correlativo del registro o c�digo �nico de la operaci�n")
    columna = columna + 1

    Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Fecha de emisi�n del comprobante de pago o documento", 11)
    columna = columna + 1
    
    If FlagFecVec Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Fecha de vencimiento y/o pago", 11)
        columna = columna + 1
    End If
    
    If TipoReg = 0 And Not FlagAnDua Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Comprobante de pago o documento")
    Else
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 2, "Comprobante de pago o documento")
    End If

    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Tipo", 3.3)
    columna = columna + 1
    
    Select Case TipoReg
        Case 1, 6
            Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "N� de Serie", 4.5)
            columna = columna + 1
            Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "N�mero")
            ColNumero = columna
            columna = columna + 1
        Case 0
            Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Serie o Cod Aduanero", 4.5)
            columna = columna + 1
            
            If FlagAnDua Then
                Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "A�o DUA", 4)
                columna = columna + 1
            End If
            
            Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "N� Doc, Formulario, DUA, DSI, Liquid. Cobranza u Otros Doc SUNAT para acreditar cr�dito fiscal")
            ColNumero = columna
            columna = columna + 1
    End Select
    
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 2, IIf(TipoReg = 1, "Informaci�n del cliente", "Informaci�n del proveedor"))
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna + 1, "Documento de Identidad")
    
    Call SetEstiloCab(o_Hoja, Fila + 2, columna, Fila + 2, columna, "Tipo", 3.3)
    columna = columna + 1
    Call SetEstiloCab(o_Hoja, Fila + 2, columna, Fila + 2, columna, "Numero", 13)
    columna = columna + 1
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Apellidos y Nombres, Denominaci�n o Raz�n Social", 50)
    ColNomEntidad = columna
    columna = columna + 1
    
    '==== Secci�n exclusiva gastos =====
    If FlagDetalle Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Detalle", 40)
        columna = columna + 1
    End If
    
    If FlagInafecto And TipoReg = 6 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Importe de la operaci�n inafecta")
        columna = columna + 1
    End If
    
    '==== Secci�n exclusiva ventas =====
    
    If FlagExport Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Valor facturado de la exportaci�n")
        columna = columna + 1
    End If
        
    If FlagAfecto Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Base Imponible de la operaci�n gravada")
        columna = columna + 1
    End If
    
    If (FlagExonerado Or FlagInafecto) And TipoReg = 1 Then
        If FlagExonerado And FlagInafecto Then
            Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Importe total de la operaci�n exonerada e inafecta")
        Else
            Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna, "Importe total de la operaci�n exonerada e inafecta")
        End If
        
        If FlagExonerado Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Exonerada")
            columna = columna + 1
        End If
        
        If FlagInafecto Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Inafecta")
            columna = columna + 1
        End If
    End If
    
    If FlagISC And TipoReg = 1 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "ISC")
        columna = columna + 1
    End If
        
    If FlagIGV Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "IGV Y/O IPM")
        columna = columna + 1
    End If
        
    If FlagBIvap Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Base imponible IVAP")
        columna = columna + 1
    End If
        
    If FlagIvap Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "IVAP")
        columna = columna + 1
    End If
    
    '==== Secci�n exclusiva compras =====
    
    If FlagMGrv1 Or FlagMIgv1 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Adquisiciones Gravadas destinadas a operaciones gravadas y/o de exportaci�n")
        
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Base Imponible")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "IGV")
        columna = columna + 1
    End If
    
    If FlagMGrv2 Or FlagMIgv2 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Adquisiciones Gravadas destinadas a operaciones gravadas y/o de exportaci�n y a operaciones no gravadas")
        
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Base Imponible")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "IGV")
        columna = columna + 1
    End If
    
    If FlagMGrv3 Or FlagMIgv3 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Adquisiciones Gravadas destinadas a operaciones no gravadas")
        
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Base Imponible")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "IGV")
        columna = columna + 1
    End If
    
    If FlagNoGrav Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Valor de adquisiciones no gravadas")
        columna = columna + 1
    End If
    
    If FlagISC And TipoReg <> 1 Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "ISC")
        columna = columna + 1
    End If
    '====================================
    
    
    If FlagPerRet Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, IIf(TipoReg = 1, "Retenci�n", "Percepci�n"))
        columna = columna + 1
    End If
        
    If FlagOtros Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Otros tributos y cargos que no forman parte de la base imponible")
        columna = columna + 1
    End If
        
    If FlagTotal Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Importe total del comprobante de pago")
        columna = columna + 1
    End If
    
    If FlagNoDom Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "N� de Comprobante de pago emitido por sujeto no domiciliado")
        columna = columna + 1
    End If
    
    If FlagDetracNum Or FlagDetracFec Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 1, "Constancia de dep�sito de detracci�n")
        
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "N�mero")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Fecha de Emisi�n")
        columna = columna + 1
    End If
    
    If FlagTipoCambio Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Tipo de cambio", 6)
        columna = columna + 1
    End If
    
    If FlagRFecha Or FlagRTipo Or FlagRSerie Or FlagRNumero Then
        Dim NumRCampos As Integer
        NumRCampos = 0
        If FlagRFecha Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna + NumRCampos, Fila + 2, columna + NumRCampos, "Fecha", 11)
            NumRCampos = NumRCampos + 1
        End If
        
        If FlagRTipo Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna + NumRCampos, Fila + 2, columna + NumRCampos, "Tipo", 3.3)
            NumRCampos = NumRCampos + 1
        End If
        
        If FlagRSerie Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna + NumRCampos, Fila + 2, columna + NumRCampos, "Serie", 4.5)
            NumRCampos = NumRCampos + 1
        End If
        
        If FlagRNumero Then
            Call SetEstiloCab(o_Hoja, Fila + 1, columna + NumRCampos, Fila + 2, columna + NumRCampos, "N� del comprobante o documento", 15)
            NumRCampos = NumRCampos + 1
        End If
        
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + NumRCampos - 1, "Fecha")
        columna = columna + NumRCampos
    End If
    
    If FlagCuenta Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 2, columna, "Cuenta")
        columna = columna + 1
    End If
    If FlagNaturaleza Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 5, "Centros de Costo")
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Naturaleza")
        columna = columna + 2
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "CCostos1")
        columna = columna + 2
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "CCostos2")
        columna = columna + 2
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "CCostos3")
        columna = columna + 2
        
    End If
    
    If FlagCCostos Then
        Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 5, "Centros de Costo")
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Admin")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Vent")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Prod")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Exclu")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Divers")
        columna = columna + 1
        Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 2, columna, "Finan")
        columna = columna + 1
    End If
    
    UltColumna = columna - 1
    Call SetEstiloGenCab(o_Hoja, Fila, 1, Fila + 2, UltColumna)

End Sub

Private Sub EscribeCabecerasKardex(ByRef o_Hoja As Object, Fila As Long, columna As Long, ByVal in_producto As String)

   
    Fila = Fila + 1
    o_Hoja.Cells(Fila, columna).Value = "ARTICULO: " & in_producto & Space(2) & "-" & get_producto(in_producto)
    
    Fila = Fila + 1
    o_Hoja.Cells(Fila, columna).Value = "U.MEDIDA: " & get_unidad_descripcion(in_producto)
    
    Fila = Fila + 1
    o_Hoja.Cells(Fila, columna).Value = "T.EXISTENCIA: " & "MERCADERIA"
    
    Fila = Fila + 1
    
  
    
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 3, "DOCUMENTO DE TRASLADO, COMPROBANTE DE PAGO, DOCUMENTO INTERNO O SIMILAR")
   
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "FECHA", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "TIPO", 6)
    columna = columna + 1
    
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "SERIE", 7)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "NUMERO", 10)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila + 1, columna, "TIPO OPERACION", 10)
    columna = columna + 1
    
    '==== CABECERA
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 2, "ENTRADAS")
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "CANTIDAD", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO.UNIT", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO TOTAL", 15)
    columna = columna + 1
    
    '==== CABECERA
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 2, "SALIDAS")
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "CANTIDAD", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO.UNIT", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO TOTAL", 15)
    columna = columna + 1
    
    '==== CABECERA
    Call SetEstiloCab(o_Hoja, Fila, columna, Fila, columna + 2, "SALDO FINAL")
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "CANTIDAD", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO.UNIT", 15)
    columna = columna + 1
    
    Call SetEstiloCab(o_Hoja, Fila + 1, columna, Fila + 1, columna, "COSTO TOTAL", 15)
    columna = columna + 1
    
    
    UltColumna = columna - 1
    Call SetEstiloGenCab(o_Hoja, Fila, 1, Fila + 1, UltColumna)
     Fila = Fila + 2

End Sub

Private Sub SetEstiloCab(ByRef o_Hoja As Object, IniFila As Long, IniColumna As Long, FinFila As Long, FinColumna As Long, NomCab As String, Optional Ancho As Double = -1)

    If Ancho >= 0 Then
        o_Hoja.Columns(IniColumna).columnWidth = Ancho
    End If
    
    With o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna))
        .MergeCells = True
        .horizontalAlignment = xlCenter
        .verticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .shrinkToFit = False
        .ReadingOrder = xlContext
        .Value = NomCab
    End With
End Sub

Private Sub SetEstiloGenCab(ByRef o_Hoja As Object, IniFila As Long, IniColumna As Long, FinFila As Long, FinColumna As Long)
    
    Select Case TipoReg
        Case 0
            o_Hoja.Rows(IniFila).RowHeight = 70
        Case 1
            o_Hoja.Rows(IniFila).RowHeight = 40
        Case 10
            o_Hoja.Rows(IniFila).RowHeight = 70
    End Select
    
    
    
    o_Hoja.Rows(IniFila + 1).RowHeight = 15
    o_Hoja.Rows(IniFila + 2).RowHeight = 20
    
    FinFilaCab = IniFila + 1
    
    With o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna)).Font
        .Bold = True
        .Size = 9
    End With
     
    Call SetBoldBorders(o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna)))
        
End Sub

Private Sub SetEstiloGenData(ByRef o_Hoja As Object, IniFila As Long, IniColumna As Long, FinFila As Long, FinColumna As Long)
    
    Call SetBoldBorders(o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna)))
    
End Sub

Private Sub SetEstiloGenTotales(ByRef o_Hoja As Object, IniFila As Long, IniColumna As Long, FinFila As Long, FinColumna As Long)
    
    With o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna)).Font
        .Bold = True
        .Size = 10
    End With
     
    Call SetBoldBorders(o_Hoja.Range(o_Hoja.Cells(IniFila, IniColumna), o_Hoja.Cells(FinFila, FinColumna)))
        
End Sub

Private Sub SetBoldBorders(ByRef Rng As Object)
    With Rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
End Sub

Private Sub optCustom_Click()
    If optCustom.Value Then
        fraCustom.Enabled = True
    Else
        fraCustom.Enabled = False
    End If
End Sub

Private Sub optSoloValores_Click()
    If optCustom.Value Then
        fraCustom.Enabled = True
    Else
        fraCustom.Enabled = False
    End If
End Sub

Private Sub optTodos_Click()
    If optCustom.Value Then
        fraCustom.Enabled = True
    Else
        fraCustom.Enabled = False
    End If
End Sub

Private Sub ResetFlags()
    FlagTotal = False
    FlagFecVec = False
    FlagAnDua = False
    FlagExport = False
    FlagAfecto = False
    FlagExonerado = False
    FlagInafecto = False
    FlagISC = False
    FlagIGV = False
    FlagBIvap = False
    FlagIvap = False
    FlagMGrv1 = False
    FlagMIgv1 = False
    FlagMGrv2 = False
    FlagMIgv2 = False
    FlagMGrv3 = False
    FlagMIgv3 = False
    FlagNoGrav = False
    FlagNoDom = False
    FlagDetracNum = False
    FlagDetracFec = False
    
    FlagDetalle = False
    FlagCuenta = False
    FlagCCostos = False
    
    FlagOtros = False
    FlagPerRet = False
    FlagTipoCambio = False
    FlagRFecha = False
    FlagRTipo = False
    FlagRSerie = False
    FlagRNumero = False
    FlagTipoOpe = False
    
    
    
    
End Sub





















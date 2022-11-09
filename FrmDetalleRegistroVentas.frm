VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleRegistroVentas 
   BorderStyle     =   0  'None
   Caption         =   "Registro Ventas"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "IMPORTACION B.DATOS"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORTACION EXCEL"
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox TxtEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox txtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtAnio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3000
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
            Picture         =   "FrmDetalleRegistroVentas.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroVentas.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcMes 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   "DataCombo1"
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5400
      TabIndex        =   9
      Top             =   2880
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
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   10
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
   Begin VB.Label Label2 
      Caption         =   "FORMAT:1401XXXXXXXXXXX"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   810
      TabIndex        =   8
      Top             =   1980
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label LblRuc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   5
      Top             =   900
      Width           =   435
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVO REGISTRO DE VENTAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   1110
      TabIndex        =   4
      Top             =   300
      Width           =   2295
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   420
      TabIndex        =   3
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   1380
      Width           =   435
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2715
      Left            =   120
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmDetalleRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call Excel_a_Access(App.Path & "\bd1.mdb", _
                    App.Path & "\Libro1.xls", _
                    "Tabla1", 10, 3)
End Sub

Private Sub Command2_Click()
Call ImportarVentas(Trim(Me.TxtRuc.Text))
End Sub
Private Sub ImportarVentas(ByVal ruc As String)
Dim rstRemoto As New ADODB.Record
  Set rstT = Nothing
  Set rstT = New ADODB.Recordset
  rstT.CursorLocation = adUseClient
  strCadena = "SELECT * FROM RegistroVentasDetalle WHERE mes='" & formato_item(Me.dtcmes.BoundText, 2) & "' AND anio='2013' AND ruc='" & ruc & "' ORDER BY  codigounico ASC"
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
    
    Documento = "-"
    horario = Format(Time, "hh:mm")
            strCadena = "P_insert_venta('" & formato_item(rstT("doc_cod"), 4) & "','00001','" & formato_item(rstT("idformapago"), 2) & "','" & formato_item(rstT("moneda"), 5) & "','no'," & _
            "'" & formato_item(Trim(rstT("serie")), 3) & "','" & formato_item(Trim(rstT("numero")), 6) & "','" & rstT("RucCliente") & "','" & rstT("NombreCliente") & "','" & rstT("afecto") & "','" & rstT("igv") & "','" & rstT("exonerado") & "','" & rstT("total") & "','0'," & _
            "'" & rstT("total") & "','0','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & rstT("tc") & "','no','" & rstT("mes") & "','" & rstT("anio") & "','" & Documento & "','" & horario & "','T','--','" & ruc & "')"
            CnBd.Execute (strCadena)
             
            
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            If IsNull(rstT("fecha_factura")) = False Then
                    strCadena = "UPDATE movimiento_venta SET fecha_fact='" & Format(rstT("fecha_factura"), "YYYY-mm-dd") & "',id_doc_fact='" & formato_item(rstT("doc_cod_factura"), 4) & "',serie_fact='" & formato_item(rstT("serie_factura"), 3) & "',numero_fact='" & formato_item(rstT("numero_factura"), 6) & "' WHERE id_venta='" & id_venta & "' AND ruc='" & ruc & "'"
                    CnBd.Execute (strCadena)
                     
                End If
            
            If rstT("anulado") = "V" Then
                strCadena = "UPDATE movimiento_venta SET anulado='si',id_cliente='" & rstT("RucCliente") & "',valor_venta='0',exonerado='0',igv='0',total='0',saldo='0' WHERE id_venta='" & id_venta & "'"
                CnBd.Execute (strCadena)
                 
            End If
            rstT.MoveNext
            DoEvents
    Next i
    MsgBox "EXITOSO"
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
Me.TxtEmpresa.Text = Trim(FrmRegistroVentas.TxtEmpresa.Text)
Me.TxtRuc.Text = Trim(FrmRegistroVentas.TxtRuc.Text)
strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM meses " & _
  " ORDER BY id_mes ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcmes)
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

Private Sub Save()
Dim descripcion As String

  If Me.TxtRuc.Text = "" Or Val(Me.txtanio.Text) < 1 Or Me.txtanio.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
      
     Select Case FrmRegistroVentas.Procedencia
     Case nuevo
          strCadena = "SELECT * FROM registro_ventas WHERE ruc='" & Trim(Me.TxtRuc.Text) & "' AND mes='" & Trim(Me.dtcmes.BoundText) & "' AND anio='" & Trim(Me.txtanio.Text) & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
                descripcion = "REGISTRO VENTAS :" + Space(5) + Me.dtcmes.Text
                strCadena = "INSERT INTO registro_ventas(ruc,mes,anio,descripcion,razon) VALUES ('" & Trim(Me.TxtRuc.Text) & "','" & Trim(Me.dtcmes.BoundText) & "','" & Trim(Me.txtanio.Text) & "','" & descripcion & "','" & Trim(Me.TxtEmpresa.Text) & "')"
                CnBd.Execute (strCadena)
                 
                Call FrmRegistroVentas.actualizar
                
                
                
                
                
            
                
                
                
                
                
                
                
                Unload Me
            Else
                MsgBox "Mes ya Registrado para dicha Empresa", vbInformation, "Mensaje para el Usuario"
          End If
          Set rst = Nothing
            
            
      Case Modificar
            
              '  StrCadena = "UPDATE Comprobantes SET doc_des='" & Me.TxtDescripcion.Text & "'," & _
                "doc_abrev='" & Me.TxtAbvreviatura.Text & "'," & _
                "cTipoMovimiento='" & Me.DtcTipoMov.BoundText & "', doc_tienda=" & _
                " '" & DocTienda & "' WHERE doc_cod= '" & Trim(Me.LblCodComprobante.Caption) & "'"
            
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
    End Select
  End If

End Sub
Private Sub Excel_a_Access(Path_BD As String, _
                           Path_XLS As String, _
                           La_Tabla As String, _
                           Filas As Integer, _
                           Columnas As Integer)
  
  
Dim Obj_Excel As Object
Dim Obj_Hoja As Object
Dim Fila_Actual As Integer
Dim Columna_Actual As Integer
Dim Dato As Variant
Dim Nombre As String
Static acum As Double
Nombre = App.Path & "\excel\" & Trim(Me.TxtRuc.Text) & "\" & "014112013" & Trim(Me.TxtRuc.Text) & ".xlsx"
  
    Screen.MousePointer = vbHourglass
  
    'Nueva instancia de Excel
    Set Obj_Excel = CreateObject("Excel.Application")
  
    ' Abre el libro de Excel
    Obj_Excel.Workbooks.Open FileName:=Nombre
  
    ' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel
    End If
      
    'Abre una nueva conexión Ado
    
      
    'Se posiciona al final    If rst_Ado.RecordCount <> 0 Then rst_Ado.MoveLast
    ' Recorre las filas y columnas de la hoja
    
        'Nuevo registro
       ' rst_Ado.AddNew
       i = 10
       Fila_Actual = i
       acum = 0
       Do While (i < 10000)
        
            
            id_doc = formato_item(Trim$(Obj_Hoja.Cells(Fila_Actual, 4)), 4)
            forma_pago = "01"
            nmoneda = "00001"
            id_delivery = "no"
            serie = formato_item(Trim$(Obj_Hoja.Cells(Fila_Actual, 5)), 3)
            nnumero = formato_item(Mid(Trim$(Val(Obj_Hoja.Cells(Fila_Actual, 6))), 1, 6), 6)
            
            id_cliente = Trim$(Obj_Hoja.Cells(Fila_Actual, 8))
            If id_cliente = "" Or id_cliente = "0" Then
                id_cliente = "00000000"
            
            End If
            
            NCLIENTE = Trim$(Obj_Hoja.Cells(Fila_Actual, 9))
            
            valor_venta = Val(Obj_Hoja.Cells(Fila_Actual, 11))
            igv = Val(Obj_Hoja.Cells(Fila_Actual, 15))
           exonerado = Val(Obj_Hoja.Cells(Fila_Actual, 12))
           
            
            
            
            tTotal = Val(Obj_Hoja.Cells(Fila_Actual, 20))
            Saldo = 0#
            monto_pago = tTotal
            Monto_Vuelto = 0#
            fechaI = Format(Obj_Hoja.Cells(Fila_Actual, 2), "YYYY-mm-dd")
            fecha_vencimiento = Format(Obj_Hoja.Cells(Fila_Actual, 3), "YYYY-mm-dd")
            
            id_tipo_factura = Trim$(Obj_Hoja.Cells(Fila_Actual, 7))
            id_vendedor = KEY_USUARIO
            tc = Format(Val(Obj_Hoja.Cells(Fila_Actual, 21)), "#,##0.00")
            afecta_factura = "no"
            id_mes = Trim(Me.dtcmes.BoundText)
            id_anio = str(Year(KEY_FECHA))
            ruc = Trim(Me.TxtRuc.Text)
        
            If Val(Obj_Hoja.Cells(Fila_Actual, 1)) > 0 Then
            
            
            horario = Format(Time, "hh:mm")
            Documento = id_doc & ":" & Trim(serie) & "-" & Trim(nnumero)
            strCadena = "P_insert_venta('" & id_doc & "','00001','" & forma_pago & "','" & nmoneda & "','" & id_delivery & "'," & _
            "'" & Trim(serie) & "','" & Trim(nnumero) & "','" & id_cliente & "','" & NCLIENTE & "','" & valor_venta & "','" & igv & "','" & exonerado & "','" & tTotal & "','0'," & _
            "'" & tTotal & "','0','" & fechaI & "','" & fecha_vencimiento & "','00001','" & KEY_USUARIO & "','" & tc & "','no','" & id_mes & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','M','direccion','" & ruc & "')"
            CnBd.Execute (strCadena)
             
            
            
            
            Else
                acum = acum + 1
                If acum > 10 Then
                    GoTo salir
                End If
            End If
        
       
        
        
        i = i + 1
        Fila_Actual = i
        DoEvents
        Loop
    
salir:
       MsgBox " Datos copiados ", vbInformation
  
Exit Sub
  

  


      
End Sub


Private Sub TxtRuc_Change()
strCadena = "SELECT * FROM persona where dni='" & Trim(Me.TxtRuc.Text) & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    
    Me.TxtEmpresa.Text = rstT("nombre_completo")
End If
End Sub

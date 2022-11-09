VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleRegistroCompras 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "IMPORTACION B.DATOS"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORTACION EXCEL"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox TxtAnio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox TxtEmpresa 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2280
      Width           =   5775
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
            Picture         =   "FrmDetalleRegistroCompras.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleRegistroCompras.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5340
      TabIndex        =   3
      Top             =   2880
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
         TabIndex        =   4
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
   Begin MSDataListLib.DataCombo DtcMes 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
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
   Begin VB.Label Label2 
      Caption         =   "FORMAT:801XXXXXXXXXXX"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   1200
      Width           =   2655
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
      Left            =   540
      TabIndex        =   10
      Top             =   1380
      Width           =   435
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
      Left            =   540
      TabIndex        =   9
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo Registro de Compras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1170
      TabIndex        =   8
      Top             =   300
      Width           =   3135
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
      Left            =   540
      TabIndex        =   7
      Top             =   900
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   7575
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
      Left            =   540
      TabIndex        =   6
      Top             =   1980
      Width           =   465
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2715
      Left            =   120
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmDetalleRegistroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Obj_Excel As Object
Dim Obj_Hoja As Object
Dim Fila_Actual As Integer
Dim Columna_Actual As Integer
Dim Dato As Variant
Dim Nombre As String
Static acum As Double
Nombre = App.Path & "\excel\" & Trim(Me.TxtRuc.Text) & "\" & "0801112013" & Trim(Me.TxtRuc.Text) & ".xlsx"
Screen.MousePointer = vbHourglass
Set Obj_Excel = CreateObject("Excel.Application")
Obj_Excel.Workbooks.Open FileName:=Nombre
  If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel
    End If
      
    
       i = 10
       Fila_Actual = i
       acum = 0
       Do While (i < 10000)
        
            
            id_doc = formato_item(Trim$(Obj_Hoja.Cells(Fila_Actual, 4)), 4)
            forma_pago = "01"
            nmoneda = "00001"
            id_delivery = "no"
            serie = formato_item(Trim$(Obj_Hoja.Cells(Fila_Actual, 5)), 3)
            nnumero = formato_item(Trim$(Obj_Hoja.Cells(Fila_Actual, 7)), 8)
            valor_venta = 0
            id_cliente = Trim$(Obj_Hoja.Cells(Fila_Actual, 9))
            dua = Format(Trim$(Obj_Hoja.Cells(Fila_Actual, 6)), "00")
            If id_cliente = "" Then
                id_cliente = "00000000"
            
            End If
            igv = 0
            valor_venta = 0
            tTotal = 0
            If id_cliente = "20131312955" Then
            d = l
            End If
            NCLIENTE = Replace(Trim$(Obj_Hoja.Cells(Fila_Actual, 10)), "N?", " ")
            
            valor_venta1 = Val(Obj_Hoja.Cells(Fila_Actual, 11))
            igv1 = Val(Obj_Hoja.Cells(Fila_Actual, 12))
            
            valor_venta2 = Val(Obj_Hoja.Cells(Fila_Actual, 13))
            igv2 = Val(Obj_Hoja.Cells(Fila_Actual, 14))
            
            valor_venta3 = Val(Obj_Hoja.Cells(Fila_Actual, 15))
            igv3 = Val(Obj_Hoja.Cells(Fila_Actual, 16))
            tipo_compra = "03"
            If Val(valor_venta1) > 0 Then
                tipo_compra = "01"
                valor_venta = valor_venta1
                igv = igv1
            End If
            If Val(valor_venta2) > 0 Then
                tipo_compra = "02"
                valor_venta = valor_venta2
                igv = igv2
            End If
            If Val(valor_venta3) > 0 Then
                tipo_compra = "03"
                valor_venta = valor_venta3
                igv = igv3
            End If
            exonerado = Format(Val(Obj_Hoja.Cells(Fila_Actual, 17)), "###0.00")
           
            
            igv = Format(Val(igv), "###0.00")
            valor_venta = Format(Val(valor_venta), "###0.00")
            tTotal = Format(Val(Obj_Hoja.Cells(Fila_Actual, 21)), "###0.00")
            Saldo = 0#
            monto_pago = tTotal
            Monto_Vuelto = 0#
            fechaI = Format(Obj_Hoja.Cells(Fila_Actual, 2), "YYYY-mm-dd")
            fecha_vencimiento = Format(Obj_Hoja.Cells(Fila_Actual, 3), "YYYY-mm-dd")
            
            id_tipo_factura = Val(Trim$(Obj_Hoja.Cells(Fila_Actual, 8)))
            id_vendedor = KEY_USUARIO
            tc = Format(Val(Obj_Hoja.Cells(Fila_Actual, 24)), "###0.00")
            afecta_factura = "no"
            id_mes = Trim(Me.dtcmes.BoundText)
            id_anio = str(Year(KEY_FECHA))
            isc = Format(Val(Obj_Hoja.Cells(Fila_Actual, 18)), "###0.00")
            percepcion = Format(Val(Obj_Hoja.Cells(Fila_Actual, 19)), "###0.00")
            ruc = Trim(Me.TxtRuc.Text)
        
            If Val(Obj_Hoja.Cells(Fila_Actual, 1)) > 0 Then
            
           NCLIENTE = Replace(NCLIENTE, "'", "")
           ' NCLIENTE = "ZONA REGISTRAL NRO SEDE MOYOBAMBA"
        'serie = "011"
        'id_doc = "0003"
        strCadena = "INSERT INTO movimiento_compra(fecha_emision,fecha_cancelacion,id_forma_pago,id_tipo_compra,anio_dua,id_moneda,id_mes,id_anio,id_alm,id_doc,serie,numero,tipo_doc_identidad,id_proveedor,nproveedor,tc," & _
        "tc_diferencia,valor_venta,igv,isc,ivap,percepcion,retencion,exonerado,otros,total,saldo,anulado,dni_save,observacion,ruc) " & _
        "VALUES('" & fechaI & "','" & fecha_vencimiento & "','02','" & tipo_compra & "','" & dua & "','" & nmoneda & "','" & formato_item(Me.dtcmes.BoundText, 2) & "','" & Year(KEY_FECHA) & "','" & KEY_ALM & "','" & id_doc & "','" & Trim(serie) & "' " & _
        ",'" & Trim(nnumero) & "','" & id_tipo_factura & "','" & Trim(id_cliente) & "','" & UCase(NCLIENTE) & "','" & tc & "','0.00','" & valor_venta & "','" & igv & "','" & isc & "','0.00','" & percepcion & "','0.00','" & exonerado & "','0.00','" & tTotal & "','" & tTotal & "','no','" & KEY_USUARIO & "','--','" & Trim(Me.TxtRuc.Text) & "')"
        CnBd.Execute (strCadena)
         
         'strCadena = "P_insert_compra('" & id_doc & "','00001','" & fechaI & "','" & fecha_vencimiento & "','02'," & _
        "'" & tipo_compra & "','" & dua & "','" & nmoneda & "','" & formato_item(Me.DtcMes.BoundText, 2) & "','" & Year(KEY_FECHA) & "','" & Trim(serie) & "'," & _
        "'" & Trim(nnumero) & "','" & id_tipo_factura & "','" & Trim(id_cliente) & "','" & UCase(NCLIENTE) & "','" & tc & "'," & _
        "'0.00','" & valor_venta & "','" & igv & "','" & isc & "','0.00','" & percepcion & "','0.00','" & exonerado & "','0.00','" & tTotal & "','" & tTotal & "','" & KEY_USUARIO & "','--','" & Trim(Me.txtRuc.text) & "')"
        'CnBd.Execute (strCadena)

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

Private Sub Command2_Click()
Call Importarcompras(Trim(Me.TxtRuc.Text))
End Sub
Private Sub Importarcompras(ByVal ruc As String)
Dim rstRemoto As New ADODB.Record
Dim valorVenta As Double
Dim igv As Double
Dim Total As Double

  Set rstT = Nothing
  Set rstT = New ADODB.Recordset
  rstT.CursorLocation = adUseClient
  strCadena = "SELECT * FROM RegistroComprasDetalle WHERE mes='" & formato_item(Me.dtcmes.BoundText, 2) & "' AND anio='2013' AND ruc='" & Trim(Me.TxtRuc.Text) & "' ORDER BY  codigounico ASC"
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
        If Len(Trim(rstT("RucCliente"))) = 8 Then
            cod_identidad = 1
        End If
        If Len(Trim(rstT("RucCliente"))) = 11 Then
            cod_identidad = 6
        End If
    
         If rstT("tipo_compra") = "01" Then
            valorVenta = rstT("grav1")
            igv = rstT("igv1")
         End If
         If rstT("tipo_compra") = "02" Then
            valorVenta = rstT("grav2")
            igv = rstT("igv2")
         End If
         If rstT("tipo_compra") = "03" Then
            valorVenta = rstT("grav3")
            igv = rstT("igv2")
         End If
         
        strCadena = "SELECT * FROM persona WHERE dni='" & rstT("RucCliente") & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "P_insert_persona('" & Trim(rstT("RucCliente")) & "','--','--','--','" & rstT("NombreCliente") & "','TARAPOTO','-','-','no','no','no','no','no','no','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
            End If
            
        
        strCadena = "P_insert_compra('" & formato_item(rstT("doc_cod"), 4) & "','00001','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','" & Format(rstT("fecha_cancelacion"), "YYYY-mm-dd") & "','02'," & _
        "'" & rstT("tipo_compra") & "','','" & formato_item(rstT("moneda"), 5) & "','" & rstT("mes") & "','" & rstT("anio") & "','" & formato_item(rstT("serie"), 3) & "'," & _
        "'" & formato_item(rstT("numero"), 8) & "','" & cod_identidad & "','" & Trim(rstT("RucCliente")) & "','" & UCase(rstT("NombreCliente")) & "','" & Val(rstT("tc")) & "'," & _
        "'0','" & valorVenta & "','" & igv & "','" & Val(rstT("isc")) & "','0','" & Val(rstT("percepcion")) & "','0','" & Val(rstT("nograv")) & "','0','" & Val(rstT("total")) & "','" & Val(rstT("total")) & "','" & KEY_USUARIO & "','--','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")

        
        strCadena = "UPDATE movimiento_compra SET anulado='si',nproveedor='A N U L A D O',id_proveedor='',valor_venta='0',igv='0',isc='0',ivap='0',percepcion='0',retencion='0',exonerado='0',otros='0',total=0,saldo='0',isc='0',percepcion='0',otros='0' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
        rstT.MoveNext
        DoEvents
    Next i
End If

  
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
Me.TxtEmpresa.Text = Trim(FrmRegistroCompras.TxtEmpresa.Text)
Me.TxtRuc.Text = Trim(FrmRegistroCompras.TxtRuc.Text)
strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM meses " & _
  " ORDER BY id_mes ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcmes)
  Me.dtcmes.BoundText = formato_item(Month(KEY_FECHA), 2)
  Me.txtanio.Text = str(Year(KEY_FECHA))
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
      
     Select Case FrmRegistroCompras.Procedencia
     Case nuevo
          strCadena = "SELECT * FROM registro_compras WHERE ruc='" & Trim(Me.TxtRuc.Text) & "' AND mes='" & Trim(Me.dtcmes.BoundText) & "' AND anio='" & Trim(Me.txtanio.Text) & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
                descripcion = "REGISTRO COMPRAS :" + Space(5) + Me.dtcmes.Text
                strCadena = "INSERT INTO registro_compras(ruc,mes,anio,descripcion,razon) VALUES ('" & Trim(Me.TxtRuc.Text) & "','" & Trim(Me.dtcmes.BoundText) & "','" & Trim(Me.txtanio.Text) & "','" & descripcion & "','" & Trim(Me.TxtEmpresa.Text) & "')"
                CnBd.Execute (strCadena)
                 
                Call FrmRegistroCompras.actualizar
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


Private Sub TxtRuc_Change()
strCadena = "SELECT * FROM persona where dni='" & Trim(Me.TxtRuc.Text) & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.TxtEmpresa.Text = rstT("nombre_completo")
End If

End Sub

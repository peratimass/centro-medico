VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPersonal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5595
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox TxtApellido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.PictureBox pbImageFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   -2040
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":0454
            Key             =   "(Huella)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":2A81
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":2DA1
            Key             =   "(Tecnico)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":65E4
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":6A38
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":6E8C
            Key             =   "(Mail)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPersonal.frx":A382
            Key             =   "(Historia)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   6465
      Left            =   18960
      TabIndex        =   3
      Top             =   1080
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   11404
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6465
      _Version        =   "6.7.9782"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   2505
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   7290
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   12859
         ButtonWidth     =   1588
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Huella"
               Key             =   "(key_huella)"
               ImageKey        =   "(Huella)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   13996
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC /DNI :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4605
      TabIndex        =   8
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   495
      TabIndex        =   7
      Top             =   420
      Width           =   765
   End
   Begin VB.Label lblAcoount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   18840
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   180
      Width           =   18495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EnumFrmCliente As EnumCliente
Public Procedencia As EnumProcede

Private Sub ChkMondoAdelantado_Click()

End Sub


Private Sub DataCombo1_Click(Area As Integer)

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    FrmDetallesParametros.Procedencia = Neutro
    Unload Me
End If

If KeyCode = 40 Then
    Me.HfdPersona.SetFocus
End If
End Sub
Public Sub actualizar()

strCadena = "SELECT * FROM view_personal  WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_completo "
Call llenarGrid(Me.HfdPersona, Me)

End Sub
Public Sub actualizar_contadores()
strCadena = "SELECT P.dni,P.nombre_completo,P.direccion,P.id_departamento FROM entidad_empresa E,persona P WHERE  E.cod_unico=P.dni AND E.id_tipo_per='00022' AND id_empresa='0' ORDER BY nombre_completo"
Call llenarGridContador(Me.HfdPersona, Me)
End Sub
Private Sub Form_Load()
 CenterForm Me
 Me.Top = 10
 
 If FrmDetallesParametros.Procedencia = buscar Then
    strCadena = "SELECT * FROM entidad_empresa WHERE id_tipo_per='00022' AND id_empresa='0'"
    Call ConfiguraRst(strCadena)
    Me.lblAcoount.Caption = str(rst.RecordCount) + Space(2) + "Registrados"
    Call actualizar_contadores
    Exit Sub
 End If
 
 
 
 Call actualizar
 
End Sub
Private Sub OptApellido_Click()
End Sub



Private Sub HfdPersona_Click()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    
  If FrmDetallesParametros.Procedencia = buscar Then
      TlbAcciones.Buttons(KEY_NEW).Enabled = False
      TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
      TlbAcciones.Buttons(KEY_DELETE).Enabled = False

      TlbAcciones.Buttons(KEY_HUELLA).Enabled = True
      Exit Sub
  End If

  If KEY_CARGO = "00004" Or KEY_CARGO = "00009" Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True

  Else
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  
    TlbAcciones.Buttons(KEY_HUELLA).Enabled = False
End If
End If
End Sub

Private Sub HfdPersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FrmVentas.Procedencia = Selecionar Then
          FrmVentas.TxtCodCliente.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
          Call FrmVentas.precionar_cliente
          FrmVentas.Procedencia = Neutro
          Unload Me
          Exit Sub
    End If
     
    
    If FrmComprasPagos.Procedencia = seleccionar_per Then
       FrmComprasPagos.Procedencia = Neutro
       FrmComprasPagos.txtBusqueda_dni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       FrmComprasPagos.lbltrabajador.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       
       Call Resalta(FrmComprasPagos.txtMonto_trabajador)
       Unload Me
       Exit Sub
    End If
    
    
    If FrmVentasPersonalizada.Procedencia = Selecionar Then
        strCadena = "SELECT * FROM persona WHERE dni='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "'"
        Call ConfiguraRst(strCadena)
        FrmVentasPersonalizada.txtRUC.Text = rst("dni")
        FrmVentasPersonalizada.txtRazonSocial.Text = rst("nombre_completo")
        FrmVentasPersonalizada.txtDireccion.Text = rst("direccion")
        
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
                If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                    FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
                Else
                    FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
                End If
        Else
            FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
        End If
        FrmVentasPersonalizada.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmDetallesParametros.Procedencia = buscar Then
        FrmDetallesParametros.TxtRucContador.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmDetallesParametros.lblRazonContador.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        
        FrmDetallesParametros.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    If FrmTransferencias.Procedencia = Selecionar Then
        FrmTransferencias.TxtRucDestino.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.TxtNombreDestino.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmTransferencias.txtdireccionfiscal.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        FrmTransferencias.txtDireccionLlegada.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        Call Resalta(FrmTransferencias.TxtRucTransporte)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmCapturaHuella.Procedencia = Selecionar Then
        FrmCapturaHuella.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmCapturaHuella.lblDni.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmCapturaHuella.lblNombre.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmCapturaHuella.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    
    
    
    
    If FrmTransferencias.Procedencia = buscar Then
        FrmTransferencias.TxtRucTransporte.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonTransporte.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        Call Resalta(FrmTransferencias.TxtMarcayPlaca)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    If FrmTransferencias.Procedencia = modificar Then
        FrmTransferencias.TxtRucChofer.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonChofer.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        strCadena = "SELECT * FROM persona WHERE dni='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst("licencia")) = True Then
            FrmTransferencias.TxtLicencia.Text = ""
        Else
            FrmTransferencias.TxtLicencia.Text = rst("licencia")
        End If
        Call Resalta(FrmTransferencias.TxtLicencia)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
     If FrmTransferencias.Procedencia = buscar Then
        FrmTransferencias.TxtRucTransporte.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonTransporte.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        Call Resalta(FrmTransferencias.TxtMarcayPlaca)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
     If FrmChequeNuevo.Procedencia = buscar Then
        FrmChequeNuevo.txtRUC.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmChequeNuevo.txtRazonSocial.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmChequeNuevo.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        FrmChequeNuevo.Procedencia = Neutro
        Call Resalta(FrmChequeNuevo.Txtcentrocosto)
        Unload Me
        Exit Sub
    End If
    
      If FrmSolicitudViaticosDeclarar.Procedencia = Selecionar Then
        FrmSolicitudViaticosDeclarar.txtRUC.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmSolicitudViaticosDeclarar.txtRazonSocial.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmSolicitudViaticosDeclarar.Procedencia = Neutro
       ' FrmSolicitudViaticosDeclarar.cmdagregar.SetFocus
        Unload Me
        Exit Sub
    End If
     
   
    
    
    If FrmDetalleAlmacen.Procedencia = Selecionar Then
       FrmDetalleAlmacen.TxtCodCliente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmDetalleAlmacen.txtEncargado.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       FrmDetalleAlmacen.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    
    
    
'    If FrmOrdenCompraDet.Procedencia = buscar Then
 '      strCadena = "SELECT cPersona,NombrePersona  FROM " & _
  '     " Persona WHERE cPersona='" & Trim(Me.HfdPersona.text) & "'"
   '    Call ConfiguraRst(strCadena)
    '    If rst.RecordCount > 0 Then
     '     FrmOrdenCompraDet.TxtcodProveedor.text = rst("cPersona")
      '    FrmOrdenCompraDet.TxtProveedor.text = rst("NombrePersona")
       '   FrmOrdenCompraDet.DtcTipoTransporte.SetFocus
        'End If
 '       FrmOrdenCompraDet.Procedencia = Neutro
  '     Unload Me
   '    Set rst = Nothing
    '   Exit Sub
    'End If
     
     
   
 
    
     
     
     'If FrmOrdenCompraDet.Procedencia = Selecionar Then
     '  strCadena = "SELECT cPersona,NombrePersona,licencia FROM " & _
     '  " Persona WHERE cPersona='" & Trim(Me.HfdPersona.text) & "'"
     '  Call ConfiguraRst(strCadena)
     '   If rst.RecordCount > 0 Then
     '     FrmOrdenCompraDet.TxtCPersona.text = rst("cPersona")
     '     FrmOrdenCompraDet.TxtChofer.text = rst("NombrePersona")
     '     FrmOrdenCompraDet.TxtLicencia.text = rst("licencia")
     '     Call Resalta(FrmOrdenCompraDet.TxtLicencia)
     '   End If
     '   FrmOrdenCompraDet.Procedencia = Neutro
     '  Unload Me
     '  Set rst = Nothing
     '  Exit Sub
    'End If
    
    If FrmCompras.Procedencia = Selecionar Then
          FrmCompras.txtRUC.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          Call FrmCompras.buscar_comprobante
          FrmCompras.TxtProveedor.Text = UCase(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
          FrmCompras.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
          Unload Me
          FrmCompras.DtTipoCompra.SetFocus
          FrmCompras.Procedencia = Neutro
          Exit Sub
    End If
     If FrmComprasGastos.Procedencia = Selecionar Then
          FrmComprasGastos.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          FrmComprasGastos.lblcliente.Caption = UCase(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
          Call Resalta(FrmComprasGastos.txtMonto)
          FrmComprasGastos.Procedencia = Neutro
          Unload Me
          Exit Sub
    End If
    
  
    
If FrmListadoFacturasCompra.Procedencia = buscar Then
       strCadena = "SELECT cPersona FROM  Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       FrmListadoFacturasCompra.TxtcodProveedor.Text = Trim(rst(0))
       Call FrmListadoFacturasCompra.llenarGrid_Proveedor
       FrmListadoFacturasCompra.Procedencia = Neutro
       Unload Me
       Set rst = Nothing
        Exit Sub
End If
If FrmBusquedaDocumentos.Procedencia = buscar Then
       FrmBusquedaDocumentos.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmBusquedaDocumentos.TxtCliente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       FrmBusquedaDocumentos.Procedencia = Neutro
       FrmBusquedaDocumentos.cmdBuscarCliente.Enabled = True
       FrmBusquedaDocumentos.cmdBuscarCliente.SetFocus
       Unload Me
       Exit Sub
End If


    
    If FrmDetalleGuia.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       
          FrmDetalleGuia.TxtCodigoEmpresaTransporte.Text = rst(0)
          FrmDetalleGuia.TxtNombreEmpresaTransporte.Text = rst(1)
          FrmDetalleGuia.TxtDireccionTransporte.Text = rst(2)
          FrmDetalleGuia.TxtRuc_Transportes.Text = rst(3)
          Unload Me
       
       Set rst = Nothing
       FrmDetalleGuia.Procedencia = Neutro
       Exit Sub
    End If
    
    If FrmreciboIngresos.Procedencia = Selecionar Then
       strCadena = "SELECT dni,nombre_completo,direccion FROM " & _
       " persona WHERE dni='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
       Call ConfiguraRst(strCadena)
       
      
          'FrmreciboIngresos.TxtCodCliente.text = rst(0)
          FrmreciboIngresos.TxtCliente.Text = rst(1)
          FrmreciboIngresos.txtDireccion.Text = rst("direccion")
          FrmreciboIngresos.txtRUC.Text = rst("dni")
          FrmreciboIngresos.TxtObservacion.Text = ""
          FrmreciboIngresos.Procedencia = Neutro
          FrmreciboIngresos.TxtMontoPago.SetFocus
          
      Set rst = Nothing
       Unload Me
       FrmreciboIngresos.Procedencia = Neutro
       
      
       Exit Sub
    End If
      
    If FrmDetalleAdelanto.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sRazonSocial,sDireccionCliente1,Per_Ruc,Observacion,MontoAdelantado FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       If Trim(rst(6)) = "N" Then
          FrmDetalleAdelanto.TxtCodCliente.Text = rst(0)
          FrmDetalleAdelanto.TxtCliente.Text = rst(1)
          FrmDetalleAdelanto.txtDireccion.Text = rst(3)
          FrmDetalleAdelanto.txtRUC.Text = rst(4)
          FrmDetalleAdelanto.txtsaldo.Text = rst(7)
        Else
          FrmDetalleAdelanto.TxtCodCliente.Text = rst(0)
          FrmDetalleAdelanto.TxtCliente.Text = rst(2)
          FrmDetalleAdelanto.txtDireccion.Text = rst(3)
          FrmDetalleAdelanto.txtRUC.Text = rst(4)
          FrmDetalleAdelanto.txtsaldo.Text = rst(7)
          
       End If
       Unload Me
       
       Set rst = Nothing
       FrmDetalleAdelanto.Procedencia = Neutro
       Exit Sub
    End If
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetallePersonal.Show
      'Call Resalta(FrmDetallePersona.TxtRuc)
      Exit Sub
    Case KEY_UPDATE
      Procedencia = modificar
      FrmDetallePersonal.Show
      Exit Sub
    Case "(Historia)"
        Procedencia = Selecionar
        FrmHistoriaClinica.Show
        Exit Sub
    Case KEY_MAIL
       Procedencia = nuevo
       frmmail.Show
       Exit Sub
     Case KEY_HUELLA
         
         FrmCapturaHuella.txtDni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
         FrmCapturaHuella.lblNombre.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
         FrmCapturaHuella.lblDni.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
         FrmCapturaHuella.Show
         Exit Sub
         
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "' "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
         strCadena = "p_delete_persona('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & KEY_RUC & "')"
         CnBd.Execute (strCadena)
        
          
         Call actualizar
         Else
            MsgBox "Imposible Eliminar a este Usuario, esta Vinculado a Movimientos"
         End If
      End If
    Case KEY_MONTO
        FrmDetalleMonto.Show
    Case KEY_TECNICO
        Procedencia = Selecionar
        FrmServiciotecnico.Show
    Case "(Salir)"
        FrmDetallesParametros.Procedencia = Neutro
        FrmDetalleAlmacen.Procedencia = Neutro
        Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
          Next
         cabecera = "DNI" & vbTab & "NOMBRE PERSONAL" & vbTab & "DIRECCION" & vbTab & "CARGO" & vbTab & "FECHA INGRESO" & vbTab & "CELULAR" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("habilitado") = "si" Then
               in_estado = "HABILITADO"
            Else
               in_estado = "INACTIVO"
            End If
            Fila = rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & rst("especialidad") & vbTab & Format(rst("fecha_ingreso"), "dd-mm-YYYY") & vbTab & rst("celular") & vbTab & in_estado
            Grilla.AddItem Fila
            If rst("habilitado") = "no" Then
                For k = 1 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
             End If
        rst.MoveNext
        Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
         TlbAcciones.Buttons(KEY_DELETE).Enabled = False
       
         TlbAcciones.Buttons(KEY_HUELLA).Enabled = False
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGridContador(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 6000
           Grilla.ColWidth(3) = 1100
           
          Next
         cabecera = "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "DIRECCION FISCAL" & vbTab & "DEPARTAMENTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             strCadena = "SELECT * FROM departamentos WHERE id_depa='" & rst("id_departamento") & "'"
             Call ConfiguraRstT(strCadena)
             If rstT.RecordCount > 0 Then
                departamento = UCase(rstT("descripcion"))
            Else
                departamento = "NO REGISTRADO"
             End If
             Fila = rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & departamento
             Grilla.AddItem Fila
            Fila = ""
        rst.MoveNext
        Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
         TlbAcciones.Buttons(KEY_NEW).Enabled = False
         TlbAcciones.Buttons(KEY_DELETE).Enabled = False
         TlbAcciones.Buttons(KEY_MAIL).Enabled = False
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
    Call Mayusculas(KeyAscii)
    If KeyAscii = 13 Then
        strCadena = "SELECT * FROM view_personal WHERE ruc='" & KEY_RUC & "' and  nombre_completo LIKE '%" & Trim(Me.TxtApellido.Text) & "%' ORDER BY nombre_completo "
        Call llenarGrid(Me.HfdPersona, Me)
    End If
End Sub



Private Sub TxtRazonSocial_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)
Dim nruc As String
If KeyAscii = 13 Then

    strCadena = "SELECT * FROM view_personal WHERE dni LIKE  '%" & Trim(Me.txtRUC.Text) & "%' and ruc='" & KEY_RUC & "' ORDER BY nombre_completo "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = nuevo
        FrmDetallePersonal.Show
        If Len(Trim(Me.txtRUC.Text)) = 8 Then
            strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRUC.Text) & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
               FrmDetallePersonal.txtRUC.Text = rstK("dni")
               Call FrmDetallePersonal.LLENA_NC(rstK("dni"))
               Exit Sub
               
            Else
                nruc = "10" & Trim(Me.txtRUC.Text)
                FrmDetallePersonal.txtRUC.Text = DigitoVerificadorRUC(Trim(nruc))
                Exit Sub
            End If
        Else
                FrmDetallePersonal.txtRUC.Text = Trim(Me.txtRUC.Text)
        End If
         Call FrmDetallePersonal.precionar
        Exit Sub
    Else
         Call llenarGrid(Me.HfdPersona, Me)
    End If

End If

End Sub

Private Sub TxtTelefono_Change()
End Sub





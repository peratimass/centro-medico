VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIFrmPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   9615
   ClientLeft      =   1725
   ClientTop       =   3915
   ClientWidth     =   10560
   Icon            =   "MDIFrmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIFrmPrincipal.frx":0ECA
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5400
      Top             =   4560
   End
   Begin VB.Timer Tiner_update 
      Enabled         =   0   'False
      Interval        =   65520
      Left            =   7680
      Top             =   3360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   9270
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1323
            MinWidth        =   1323
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10560
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":501E3
            Key             =   "(venta)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":532C3
            Key             =   "(general)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":563F6
            Key             =   "(especialista)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":59658
            Key             =   "(ticket)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":5BC0A
            Key             =   "(fua)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":5DF01
            Key             =   "(procedimiento)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":61409
            Key             =   "(Notificacion)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":64614
            Key             =   "(laboratorio)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":6799C
            Key             =   "(imagenes)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":6AABF
            Key             =   "(agenda)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":6DBE6
            Key             =   "(enfermera)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":70FF1
            Key             =   "(medico)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":743CE
            Key             =   "(seguro)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":76755
            Key             =   "(agenda_medico)"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":7A431
            Key             =   "(archivo)"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":7E40C
            Key             =   "(medico1)"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":81672
            Key             =   "(triaje)"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":8540D
            Key             =   "(triaje_emergencia)"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":895B5
            Key             =   "(Compra)"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":8C5DC
            Key             =   "(Transferencia)"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":8F699
            Key             =   "(Celular)"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":924BA
            Key             =   "(Exit)"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":955F1
            Key             =   "(paciente)"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":98B66
            Key             =   "(medicamento)"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmPrincipal.frx":9B2E8
            Key             =   "(Caja)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   794
      ButtonWidth     =   820
      ButtonHeight    =   794
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   31
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes / Proveedores"
            ImageKey        =   "(paciente)"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Personal/Trabajadores"
            ImageKey        =   "(especialista)"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ventas"
            ImageKey        =   "(venta)"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Compras"
            ImageKey        =   "(Compra)"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guias Remision"
            ImageKey        =   "(Transferencia)"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Productos"
            ImageKey        =   "(general)"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Documentos Realizados"
            ImageKey        =   "(agenda)"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Permisos / Asistencia"
            ImageKey        =   "(agenda_medico)"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Caja"
            ImageKey        =   "(Caja)"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Central Telefonica"
            ImageKey        =   "(Celular)"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Produccion"
            ImageKey        =   "(archivo)"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tramite Documentario"
            ImageKey        =   "(triaje)"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Servicio Tecnico"
            ImageKey        =   "(seguro)"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Formatos de Atencion(FUA)"
            ImageKey        =   "(fua)"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Ticket Atencion"
            ImageKey        =   "(ticket)"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageKey        =   "(Exit)"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuMantenimientos 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnucliente 
         Caption         =   "Clientes / Proveedores                                                       "
      End
      Begin VB.Menu mnulineadoble 
         Caption         =   "-"
      End
      Begin VB.Menu Mnupersonal 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnuempelados 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproducto 
         Caption         =   "Productos"
         Begin VB.Menu MnuProductosListado 
            Caption         =   "&Listado Productos"
         End
         Begin VB.Menu mnulineas4hs 
            Caption         =   "-"
         End
         Begin VB.Menu mnuproductosproveedor1 
            Caption         =   "&Productos Proveedor"
         End
         Begin VB.Menu mnulinea1020304050 
            Caption         =   "-"
         End
         Begin VB.Menu mnuproductosmermas 
            Caption         =   "&Productos Mermas"
         End
         Begin VB.Menu mnulineasr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuproductocombo 
            Caption         =   "&Ensamblado"
         End
         Begin VB.Menu mnulines41 
            Caption         =   "-"
         End
         Begin VB.Menu mnuproductosDeribados 
            Caption         =   "&Productos Derivados"
         End
         Begin VB.Menu mnulineyeegege 
            Caption         =   "-"
         End
         Begin VB.Menu mnuproductostransformaciones 
            Caption         =   "&Producto Transformaciones"
         End
         Begin VB.Menu mnulineas4545ss 
            Caption         =   "-"
         End
         Begin VB.Menu mnucambiosprecios 
            Caption         =   "&Cambios de Precio"
         End
         Begin VB.Menu mnulineasublineas10 
            Caption         =   "-"
         End
         Begin VB.Menu MnuColores 
            Caption         =   "&Colores"
         End
         Begin VB.Menu mnulinecolores 
            Caption         =   "-"
         End
         Begin VB.Menu mnulineasProduccion 
            Caption         =   "&Lineas de Produccion"
         End
         Begin VB.Menu mnulineatipoproducto 
            Caption         =   "-"
         End
         Begin VB.Menu mnueditorial 
            Caption         =   "Editorial"
         End
         Begin VB.Menu mnueditoriallinea 
            Caption         =   "-"
         End
         Begin VB.Menu mnutipoproducto 
            Caption         =   "&Tipos de Producto"
         End
         Begin VB.Menu mnutanques 
            Caption         =   "Cisternas- Islas-Surtidores"
         End
      End
      Begin VB.Menu mnulineas14568741 
         Caption         =   "-"
      End
      Begin VB.Menu mnulinea 
         Caption         =   "Lineas/Clasificacion"
      End
      Begin VB.Menu mnusublineas1010 
         Caption         =   "-"
      End
      Begin VB.Menu mnusublineas 
         Caption         =   "Sub-Lineas"
      End
      Begin VB.Menu menulineamodelo 
         Caption         =   "-"
      End
      Begin VB.Menu mnumodelos 
         Caption         =   "Modelos "
      End
      Begin VB.Menu linea05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnidad 
         Caption         =   "Unidades Medida"
      End
      Begin VB.Menu Linea06 
         Caption         =   "-"
      End
      Begin VB.Menu mnulinea008 
         Caption         =   "Marcas"
      End
      Begin VB.Menu Linea08 
         Caption         =   "-"
      End
      Begin VB.Menu mnuturnos00 
         Caption         =   "Turnos de Trabajo"
      End
      Begin VB.Menu mnuturnos 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDistribuidoras 
         Caption         =   "Sucursales/Almacenes"
      End
      Begin VB.Menu Linea09 
         Caption         =   "-"
      End
      Begin VB.Menu MnuComprobantes 
         Caption         =   "Comprobantes"
      End
      Begin VB.Menu Linea010 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlancontable 
         Caption         =   "Formas Pago"
      End
      Begin VB.Menu mnulinea4579510 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinfraestructura 
         Caption         =   "Infraestrutura Empresa"
      End
      Begin VB.Menu mnulinea0045 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPlandecuentas 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu mnuuulinea004 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoCambio 
         Caption         =   "Tipo Cambio"
      End
      Begin VB.Menu mnulineas12457 
         Caption         =   "-"
      End
      Begin VB.Menu MnuZonas 
         Caption         =   "Zonas"
      End
      Begin VB.Menu mnulineas4545454 
         Caption         =   "-"
      End
      Begin VB.Menu menumermas01 
         Caption         =   "Mermas"
      End
      Begin VB.Menu mnulineasa78 
         Caption         =   "-"
      End
      Begin VB.Menu mnucargospersonal 
         Caption         =   "Cargos Personal"
      End
      Begin VB.Menu mnulinea102154 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconecptosgastos 
         Caption         =   "Planes de Servicio"
      End
      Begin VB.Menu mnulineas010201 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventariow 
         Caption         =   "Inventario"
      End
      Begin VB.Menu mnulinea14785963 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClasificacionUnidadesTranspo 
         Caption         =   "&Clasificacion Unidades Transporte"
      End
      Begin VB.Menu mnukaskajskjask 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnidadesTransporte 
         Caption         =   "&Unidades Transporte"
      End
      Begin VB.Menu mnulineas478s787s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuparametrosigv 
         Caption         =   "Parametros IGV"
      End
      Begin VB.Menu mnulineasjkajs 
         Caption         =   "-"
      End
      Begin VB.Menu mnucategoria1 
         Caption         =   "Categorias "
      End
      Begin VB.Menu mnucategoriass 
         Caption         =   "-"
      End
      Begin VB.Menu mnugradosestudio 
         Caption         =   "Gestion Educativa"
      End
      Begin VB.Menu mnulineagrado 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaseguradoras 
         Caption         =   "Aseguradoras"
      End
      Begin VB.Menu mnulineaaseguradoras 
         Caption         =   "-"
      End
      Begin VB.Menu MnuParametrosEmpresa 
         Caption         =   "Parametros Empresa"
      End
      Begin VB.Menu mnulineasparametros 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuMovimientos 
      Caption         =   "M&ovimientos"
      Begin VB.Menu MnuDocumentoVenta 
         Caption         =   "Ventas                                                           "
      End
      Begin VB.Menu mnulinea789 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCompras 
         Caption         =   "Compras"
      End
      Begin VB.Menu mnulinea10203045 
         Caption         =   "-"
      End
      Begin VB.Menu mnutracking 
         Caption         =   "Tracking"
      End
      Begin VB.Menu mnulineatracking 
         Caption         =   "-"
      End
      Begin VB.Menu mnugeneracionmenuslaidad 
         Caption         =   "Generacion Mensualidad"
      End
      Begin VB.Menu mnulineareciboas 
         Caption         =   "-"
      End
      Begin VB.Menu mnutransferencias10 
         Caption         =   "Guias de Remision"
      End
      Begin VB.Menu mnulinesjhst7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuordencompra 
         Caption         =   "&Pedidos"
      End
      Begin VB.Menu Mnulineastgsfs78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuordencompra10 
         Caption         =   "Orden de Compra"
      End
      Begin VB.Menu mnulinea10203040 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParteDiaria 
         Caption         =   "Solicitud Creditos"
      End
      Begin VB.Menu mnulineapartediaria 
         Caption         =   "-"
      End
      Begin VB.Menu mnupartematerial 
         Caption         =   "Orden Salida"
      End
      Begin VB.Menu mnulinea1024587 
         Caption         =   "-"
      End
      Begin VB.Menu mnucontrolpartediaria 
         Caption         =   "Control de Partes Diaria"
      End
      Begin VB.Menu mnulineacontrolpartediaria 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPagoFacturas 
         Caption         =   "Cuentas x Pagar"
      End
      Begin VB.Menu mnulinead47s 
         Caption         =   "-"
      End
      Begin VB.Menu MnuListadoDeudores01 
         Caption         =   "Cuentas x Cobrar"
      End
      Begin VB.Menu mnulineamisproyectos 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproyectos 
         Caption         =   "Mis Proyectos"
      End
      Begin VB.Menu mnulineacambio 
         Caption         =   "-"
      End
      Begin VB.Menu mnucambioaceite 
         Caption         =   "Servicio Tecnico"
      End
      Begin VB.Menu mnulineboni 
         Caption         =   "-"
      End
      Begin VB.Menu mnubonificaciones 
         Caption         =   "Bonificaciones"
      End
      Begin VB.Menu mnulineamemorandun 
         Caption         =   "-"
      End
      Begin VB.Menu mnumemorandum 
         Caption         =   "Memorandum"
      End
   End
   Begin VB.Menu mnucaja 
      Caption         =   "&Caja"
      Begin VB.Menu MnuSalidadeDinero 
         Caption         =   "Prestamo Personal"
      End
      Begin VB.Menu linea019 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIngresoDinero 
         Caption         =   "Recibo de Ingresos"
      End
      Begin VB.Menu mnulimo102030 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSolicitudViaticos 
         Caption         =   "Solicitud Viaticos"
      End
      Begin VB.Menu mnulinea1020l 
         Caption         =   "-"
      End
      Begin VB.Menu mnuordenpago 
         Caption         =   "&Orden de Pago"
      End
      Begin VB.Menu mnulinea45748 
         Caption         =   "-"
      End
      Begin VB.Menu mnumiscuentas 
         Caption         =   "Mis Cuentas"
      End
      Begin VB.Menu MNULINEAS0101ss 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchequeras 
         Caption         =   "Chequeras"
      End
      Begin VB.Menu mnu1024ss 
         Caption         =   "-"
      End
      Begin VB.Menu mnugastos 
         Caption         =   "Gastos"
      End
   End
   Begin VB.Menu MnuActualizacion 
      Caption         =   "&Actualizacion"
      Begin VB.Menu MnuActualizarPrecio 
         Caption         =   "Actualizar Precios                "
      End
      Begin VB.Menu mnulinea4152 
         Caption         =   "-"
      End
      Begin VB.Menu mnulinea78945 
         Caption         =   "&Kardex por Producto"
      End
      Begin VB.Menu mnulinea0147 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBusquedaDocumentos 
         Caption         =   "&Busqueda de Documentos"
      End
      Begin VB.Menu mnulinea4568 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPendientes 
         Caption         =   "Movimientos Pendientes"
      End
   End
   Begin VB.Menu MnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnureportegeneral 
         Caption         =   "Reportes Generales VII       "
      End
      Begin VB.Menu mnulinea5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRptProductos 
         Caption         =   "Reportes Generales"
      End
      Begin VB.Menu LineaProveedores 
         Caption         =   "-"
      End
      Begin VB.Menu mnuimpresioncodBarra 
         Caption         =   "Impresion Codigo Barra"
      End
      Begin VB.Menu mnulinecodigovbarra 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVentadeProductos 
         Caption         =   "Listado Comprobantes venta"
      End
      Begin VB.Menu mnulinea1245023 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRotacion 
         Caption         =   "Rotacion de Existencias"
      End
      Begin VB.Menu MnuLineValorzado414 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConsolidadoMovAlmacen 
         Caption         =   "Consolidado Movimientos de Almacen"
      End
      Begin VB.Menu Mnulineaconsoliadod 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIngresosProductos 
         Caption         =   "Ingresos y Salidas de Productos"
      End
      Begin VB.Menu MnulIneaIngreso 
         Caption         =   "-"
      End
      Begin VB.Menu mnucobranzadeldiae 
         Caption         =   "Cobranza del Día"
      End
      Begin VB.Menu mnucobranzadeldia 
         Caption         =   "-"
      End
      Begin VB.Menu MnuUtilidadesDiarios 
         Caption         =   "Utilidades Diarios"
      End
      Begin VB.Menu mnulineassjahsas 
         Caption         =   "-"
      End
      Begin VB.Menu menuReporteGastos 
         Caption         =   "Resumen Linea producto [MARGEN BRUTO]"
      End
      Begin VB.Menu mnulineas102010 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvaraicion 
         Caption         =   "Variacion de Costes de Productos"
      End
      Begin VB.Menu mnulineas1478527878 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproductosproveedor 
         Caption         =   "Productos Proveedor"
      End
      Begin VB.Menu mnulk14521 
         Caption         =   "-"
      End
      Begin VB.Menu mnuentradasalida01 
         Caption         =   "Entrada Salida Personal"
      End
      Begin VB.Menu ascf10201 
         Caption         =   "-"
      End
      Begin VB.Menu mnugastoscomercializacion 
         Caption         =   "Gastos Comercializacion"
         Begin VB.Menu mnusolicitudespersonal 
            Caption         =   "Solicitudes Personal"
         End
         Begin VB.Menu mnusalidaDinero 
            Caption         =   "Detalle Salidas Dinero"
         End
      End
      Begin VB.Menu mnulineacierres 
         Caption         =   "-"
      End
      Begin VB.Menu mnucierresucursal 
         Caption         =   "Cierres Sucursal"
      End
   End
   Begin VB.Menu MnuInformesContables 
      Caption         =   "Informes Contables"
      Begin VB.Menu mnulibrodiario 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu mnulibrodiadiosimplificado 
         Caption         =   "Libro Diario Simplificado"
      End
      Begin VB.Menu mnuLibroMayor 
         Caption         =   "Libro Mayor"
      End
      Begin VB.Menu mnuLibroCajabanco 
         Caption         =   "Almacen Caja"
      End
      Begin VB.Menu mnuLineas104741 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroCompras1 
         Caption         =   "Registro de Compras"
      End
      Begin VB.Menu mnuRegistroVentas1 
         Caption         =   "Registro Ventas"
      End
      Begin VB.Menu mnulibroHonorarios 
         Caption         =   "Libro de Honorarios"
      End
      Begin VB.Menu mnulibroretenciones 
         Caption         =   "Libro Retenciones"
      End
      Begin VB.Menu mnuRegistroretencionesproveedor 
         Caption         =   "Registros Retenciones/Proveedor"
      End
      Begin VB.Menu mnuConsiliacionbancaria 
         Caption         =   "Consiliacion Bancaria"
      End
      Begin VB.Menu mnulinea123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBalansasumasysaldos 
         Caption         =   "Balanace de Comprobancion de Sumas y Saldos"
      End
      Begin VB.Menu mnuhojatrabajao10columnas 
         Caption         =   "Hoja de Trabajo 10 Columnas "
      End
      Begin VB.Menu mnubalanceresultadosratiosfinancieros 
         Caption         =   "Balance, Estado de Resultados y Ratios Financieros"
      End
      Begin VB.Menu mnulinea123456 
         Caption         =   "-"
      End
      Begin VB.Menu mnuflujodeefctivo 
         Caption         =   "Flujo de Efectivo"
      End
      Begin VB.Menu mnulinea123457 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnalisisdocumentos 
         Caption         =   "Analisis por Cuenta [Detallado]"
      End
   End
   Begin VB.Menu mnuInformessunat 
      Caption         =   "Informes SUNAT"
      Begin VB.Menu mnulibrodiariosunat 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu mnulibrodiariosimplicado 
         Caption         =   "Libro Diario Simplificado"
      End
      Begin VB.Menu mnuLibromatorsunat 
         Caption         =   "Libro mayor"
      End
      Begin VB.Menu mnuLibroComprassunat 
         Caption         =   "Libro Compras"
      End
      Begin VB.Menu mnulibroventasunat 
         Caption         =   "Libro Ventas"
      End
      Begin VB.Menu mnuinventarioybalancesunat 
         Caption         =   "Inventario y Balance"
      End
      Begin VB.Menu mnucajaybancossunat 
         Caption         =   "10 Caja y Bancos"
      End
      Begin VB.Menu mnu101Caja 
         Caption         =   "101 Caja"
      End
      Begin VB.Menu mnu104bancos 
         Caption         =   "104 Bancos"
      End
      Begin VB.Menu mnu12CLientes 
         Caption         =   "12 Clientes"
      End
      Begin VB.Menu mnu13ctasxcorarrelacionadas 
         Caption         =   "13 Ctas x Cobrar Relacionadas"
      End
      Begin VB.Menu mnu14ctasxcobrarAccPersona 
         Caption         =   "14 Ctas x Cobrar Acc.Personal"
      End
      Begin VB.Menu mnuctasxcobrardiversas 
         Caption         =   "16 Ctas x Cobrar Diversas"
      End
      Begin VB.Menu mnu19provctasxcobrardudosa 
         Caption         =   "19 Prov. Cuentas x Cobrar Dudosa"
      End
      Begin VB.Menu mnu40tributosxpagar 
         Caption         =   "40 Tributos x Pgara"
      End
      Begin VB.Menu mnu42Proveedores 
         Caption         =   "42 Proveedores"
      End
      Begin VB.Menu mnu43cuentasxpgarrelacionadas 
         Caption         =   "43 Cuentas x Pagar Relacionadas"
      End
      Begin VB.Menu mnu46cuentasxpagardiversas 
         Caption         =   "46 Cuentas x Pagar Diversas"
      End
      Begin VB.Menu mnubeneficiossocialesdelostrabajadores 
         Caption         =   "47 Beneficios Sociales de los Trabajadores"
      End
      Begin VB.Menu mnu49GananciasDifridas 
         Caption         =   "49 Ganancias DIferidas"
      End
      Begin VB.Menu mnulineas123458 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexportaralpdbcompras 
         Caption         =   "Exportar al PDB Compras"
      End
      Begin VB.Menu mnuExportaraPDBventas 
         Caption         =   "Exportar al PDB Ventas"
      End
      Begin VB.Menu mnuexportaralPDBtc 
         Caption         =   "Exportar al PDB TC"
      End
      Begin VB.Menu mnulineas1234569 
         Caption         =   "-"
      End
      Begin VB.Menu mnudaotcompra 
         Caption         =   "DAOT Compra"
      End
      Begin VB.Menu mnudaotventa 
         Caption         =   "DAOT Venta"
      End
      Begin VB.Menu mnilinea12345612 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLibroElectronico 
         Caption         =   "Libro Electronico"
      End
   End
   Begin VB.Menu mnuInformesGerenciales 
      Caption         =   "Informes Gerenciales"
      Begin VB.Menu mniIDG 
         Caption         =   "I.D.G"
      End
      Begin VB.Menu mnuBancoGerenciales 
         Caption         =   "Bancos"
      End
      Begin VB.Menu mnulineabancosgerenciales 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCLientesgerenciales 
         Caption         =   "Clientes"
      End
      Begin VB.Menu MnuProveedoresGerenciales 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu mnuunidadesnegocio 
         Caption         =   "Unidades de Negocio"
      End
      Begin VB.Menu mnuPresupuestoGerencial 
         Caption         =   "Presupuesto"
      End
      Begin VB.Menu mnulineas147852 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBalanceGerencial 
         Caption         =   "Balance"
      End
   End
   Begin VB.Menu MnuGestionFinanciera 
      Caption         =   "Gestion Financiera"
      Begin VB.Menu mnuSaldoBancos 
         Caption         =   "Saldos de Bancos"
      End
      Begin VB.Menu mnuFlujodeCaja 
         Caption         =   "Flujo de Caja"
      End
      Begin VB.Menu mnuclientesproveedores 
         Caption         =   "Cuentas Corrientes,Clientes Proveedores"
      End
      Begin VB.Menu mnuEstadiCuentabancosfacturasletras 
         Caption         =   "Estado de Cuenta (Bancos,Facturas,Letras)"
      End
      Begin VB.Menu mniCuentasvencidasyvigentes 
         Caption         =   "Cuentas Vencidas y Vigentes"
      End
      Begin VB.Menu mnulinea147852 
         Caption         =   "-"
      End
      Begin VB.Menu mnumantecuentasxpgara 
         Caption         =   "Mantenimientos Cuentas x Pagar"
      End
      Begin VB.Menu mnuChequeVoucher 
         Caption         =   "Cheque Voucher"
      End
      Begin VB.Menu mnuChequesGirados 
         Caption         =   "Cheques Girados"
      End
      Begin VB.Menu mnulineas852369 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCajaEgreso 
         Caption         =   "Ingreso/ Egreso Caja"
      End
      Begin VB.Menu mnumantecuentasxcobrar 
         Caption         =   "Mantenimiento de Cuentas x Cobrar"
      End
      Begin VB.Menu mnucobranza 
         Caption         =   "Cobranza"
      End
      Begin VB.Menu mnuhojadecobranza 
         Caption         =   "Hoja de Cobranza"
      End
      Begin VB.Menu mnuavisovencimientoletras 
         Caption         =   "Aviso Vencimiento Letras"
      End
   End
   Begin VB.Menu mnuGestionegocio 
      Caption         =   "Gestion de Negocio"
      Begin VB.Menu mnuunidadesnegocios 
         Caption         =   "Unidades de Negocio (C.Costo)"
      End
      Begin VB.Menu mnuClasificarCuentas 
         Caption         =   "Clasificar Cuentas"
      End
      Begin VB.Menu mnulineas451023 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuSeguridad 
      Caption         =   "&Contactenos"
   End
End
Attribute VB_Name = "MDIFrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public EnumBuscar As EnumBuscarDocumento
Private Sub frmtransformaciones_Click()
    FrmProductoTransformaciones.Show
  
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = salir

If Cancel = vbYes Then
       strCadena = "UPDATE almacen SET dni_save='0' WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_VENTANILLA & "' and id_sucursal='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
    
       strCadena = "DELETE FROM gig_usuarios_online WHERE id_gigane='" & KEY_USUARIO & "'"
       CnBd.Execute (strCadena)
End


End If
End Sub
Private Sub MnuAyuda_Click()
End Sub
Private Sub MnuAccesoUsuarios_Click()
FrmListadoUsuarios.Show
End Sub

Private Sub menumermas01_Click()
FrmMermas.Show
End Sub

Private Sub menuReporteGastos_Click()
frmReporteMargenBruto.Show
End Sub

Private Sub MnuActualizarPrecio_Click()
FrmPrecios.Show
End Sub
Private Sub MnuActualziarStok_Click()
FrmActualizacion.Show
End Sub

Private Sub MnuAdelantoCliente_Click()

End Sub

Private Sub MnuAgrupacionCuentas_Click()
FrmAgrupacionCuentas.Show
End Sub

Private Sub mnuAnalisisdocumentos_Click()
frmAnalisisporCuenta.Show

End Sub

Private Sub mnuaseguradoras_Click()
FrmSeguros.Show
End Sub

Private Sub mnubonificaciones_Click()
FrmBonificacion.Show
End Sub

Private Sub MnuBusquedaDocumentos_Click()
FrmBusquedaDocumentos.Show
End Sub

Private Sub mnuCajaEgreso_Click()
frmCajaEgreso.Show
End Sub

Private Sub mnucambioaceite_Click()
FrmCambioAceite.Show
End Sub

Private Sub mnucambiosprecios_Click()
frmVariacionCostes.Show
End Sub

Private Sub mnucargospersonal_Click()
FrmCargosPersonal.Show
End Sub

Private Sub mnucategoria1_Click()
FrmCategoria.Show
End Sub

Private Sub mnuchequeras_Click()
FrmChequeras.Show
End Sub

Private Sub mnuchoripan_Click()
FrmChoripan.Show
End Sub

Private Sub mnucierresucursal_Click()
frmcierre.Show
End Sub

Private Sub mnuClasificacionUnidadesTranspo_Click()
FrmUnidadesTransporteTipo.Show
End Sub

Private Sub MnuCliente_Click()
If KEY_RUBRO = "00025" Then
   FrmMatricula.Show
Else
    FrmPersona.Show
End If
End Sub
Private Sub MnuDistribuidoraLaboratorio_Click()
  FrmDistribuidoraProveedor.Show
End Sub

Private Sub mnuclientescontabilidad_Click()
FrmClientes.Show
End Sub

Private Sub mnucobranzadeldiae_Click()
FrmReporteRecaudacionDiaria.Show
End Sub

Private Sub mnucombos0101_Click()
frmCombo.Show
End Sub

Private Sub MnuColores_Click()
FrmColores.Show
End Sub

Private Sub mnucompras_Click()
FrmCompras.Show
End Sub
Private Sub MnuComprobantes_Click()
FrmComprobantes.Show
End Sub

Private Sub mnuconecptosgastos_Click()

frmPlanesServicio.Show
End Sub

Private Sub MnuConsolidadoMovAlmacen_Click()
FrmReporteConsolidado.Show
End Sub

Private Sub MnuControlAccesos_Click()
FrmControlAccesos.Show
End Sub

Private Sub MnuDistribuidoraProveedor_Click()
FrmDistribuidoraProveedor.Show
End Sub

Private Sub MnuConversionBoletas_Click()
FrmConversionBoletas.Show
End Sub

Private Sub MnuCuentacorriente_Click()
FrmCuentasCorrientes.Show
End Sub

Private Sub MnuDetalleAdelantos_Click()
FrmDetalleAdelanto.Show
End Sub

Private Sub mnucontratos_Click()
FrmFichaIncripcion.Show
End Sub

Private Sub mnucontroldepersonal_Click()


FrmHuellaDigital.Show
End Sub

Private Sub mnuderivados_Click()
FrmDerivados.Show
End Sub

Private Sub MnuDetalleAdelantado_Click()
FrmDetalleAdelanto.Show
End Sub

Private Sub MnuDeudores_Click()
FrmPagoCredito.Show
End Sub

Private Sub mnucontrolpartediaria_Click()
FrmParteControl.Show
End Sub

Private Sub mnudaotcompra_Click()
FrmDaotCompra.Show
End Sub

Private Sub mnudaotventa_Click()
FrmDaotVenta.Show
End Sub

Private Sub MnuDistribuidoras_Click()
  FrmAlmacenes.Show
End Sub

Private Sub MnuDocumentoVenta_Click()
   'EnumBuscar = BDocumentoVenta
    FrmVentas.Show
End Sub

Private Sub MnuFacturaCompra_Click()
  EnumBuscar = BDocumentoCompra
  
End Sub

Private Sub MnuLaboratorios_Click()
  FrmProveedor.Show
End Sub

Private Sub mnueditorial_Click()
FrmEditorial.Show
End Sub

Private Sub mnuentradasalida01_Click()
FrmVigilante1.Show
End Sub

Private Sub mnugastos_Click()
frmNuevoComprobante.Show
End Sub

Private Sub mnuGastos1020_Click()
FrmProductosGastos.Show
End Sub

Private Sub mnugeneracionmenuslaidad_Click()
frmgeneracionmensualidad.Show
End Sub

Private Sub mnugradosestudio_Click()
frmgestion_college.Show
End Sub

Private Sub mnuimpresioncodBarra_Click()
'Form1.Show
FrmGeneradorBarras.Show
End Sub

Private Sub mnuinfraestructura_Click()
frmHotelInfraestructura.Show

End Sub

Private Sub MnuIngresoDinero_Click()

FrmreciboIngresos.Show
End Sub

Private Sub MnuIngresosProductos_Click()
FrmReporteIngresos.Show
End Sub

Private Sub mnuInventariow_Click()
FrmInventario.Show
End Sub

Private Sub mnuLibroCajabanco_Click()
frmComprobantesCaja.Show
End Sub

Private Sub mnulibrodiario_Click()
FrmRegistroDiario.Show
End Sub

Private Sub mnuLibroMayor_Click()
FrmRegistroMayor.Show
End Sub

Private Sub MnuLibrosAuxiliares_Click()
FrmLibrosAuxiliares.Show
End Sub

Private Sub MnuLinea_Click()
FrmLinea.Show
End Sub

Private Sub MnuLoteProducto_Click()
  
End Sub

Private Sub MnuOtraEntrada_Click()
  EnumBuscar = BOtraEntrada
  
End Sub

Private Sub MnuOtraSalida_Click()
  EnumBuscar = BOtraSalida
 
End Sub

Private Sub MnuPedidos_Click()
  EnumBuscar = BPedido
 
End Sub

Private Sub mnulinea008_Click()
FrmMarcas.Show
End Sub

Private Sub mnulinea475698_Click()

End Sub

Private Sub mnulinea78945_Click()
FrmKardexdeProductos.Show
End Sub

Private Sub MnuLineas45as4_Click()
FrmCentroCostos.Show
End Sub

Private Sub mnulineasProduccion_Click()
FrmLineasProduccion.Show
Exit Sub
End Sub

Private Sub MnuListadoDeudores01_Click()
FrmReporteRegistroVentas.Show
'strCadena = "SELECT * FROM view_listado_comprobanteii WHERE ruc='" & KEY_RUC & "' and id_doc<>'0099' and saldo>0 LIMIT 28"
'strCadena = "SELECT * FROM view_listado_comprobante_ultimate WHERE ruc='" & KEY_RUC & "' and id_doc<>'0099' and saldo>0 LIMIT 28"
'Call FrmReporteRegistroVentas.llenar_grid(FrmReporteRegistroVentas.HfdPersona)
'strCadenaII = "SELECT * FROM view_listado_comprobante_iii WHERE ruc='" & KEY_RUC & "' and saldo>0 "
    
End Sub

Private Sub mnumermas_Click()
  FrmProductoMermas.Show
End Sub

Private Sub mnumemorandum_Click()
 frmMemorandun.Show
End Sub

Private Sub mnumiscuentas_Click()
FrmMiscuentas.Show
End Sub

Private Sub mnuordendepago_Click()
FrmOrdenpago.Show
End Sub

Private Sub mnumovimientoplanta_Click()
FrmPlanta.Show
End Sub

Private Sub mnumodelos_Click()
FrmModelo.Show
End Sub

Private Sub mnuOrdenCompra_Click()
FrmPedido.Show
End Sub

Private Sub mnuordencompra10_Click()
FrmOrdenCompra.Show
End Sub

Private Sub MnuPagoFacturas_Click()
Procedencia = buscar
FrmListadoFacturasCompra.Show

End Sub

Private Sub MnuParametrosEmpresa_Click()
FrmParametrosEmpresa.Show
End Sub

Private Sub mnuparametrosigv_Click()
frmParametros.Show
End Sub

Private Sub mnuParteDiaria_Click()

FrmSolicitudCredito.Show
'FrmParteDiaria.Show
End Sub

Private Sub mnupartematerial_Click()
frmOrdeneSalida.Show
End Sub

Private Sub mnuPendientes_Click()
'FrmPendiente.Show
End Sub

Private Sub Mnupersonal_Click()
frmpersonal.Show
End Sub

Private Sub mnuPlancontable_Click()
FrmFormaPago.Show
End Sub

Private Sub mnuPlancontableCuentas_Click()
FrmPlanContableCuentas.Show
End Sub
Private Sub MnuProveedor_Click()

End Sub

Private Sub MnuPlandecuentas_Click()
FrmPlanContableCuentas.Show
End Sub

Private Sub mnuproductocombo_Click()

frmCombo.Show

End Sub

Private Sub mnuproductosDeribados_Click()
FrmDerivados.Show
End Sub
Private Sub MnuProductosListado_Click()
  
  FrmProducto.Show
  Exit Sub
  
End Sub
Private Sub mnuproductosmermas_Click()
  FrmProductoMermas.Show
End Sub

Private Sub mnuProductosProveedor_Click()
strCadena = "SELECT     Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, Producto.PrecioVenta, Producto.PrecioCompra, " & _
"Producto_Proveedor.cPersona , Persona.NombrePersona FROM         Producto_Proveedor INNER JOIN " & _
"                      Producto ON Producto_Proveedor.cProducto = Producto.cProducto INNER JOIN " & _
"                      Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN " & _
"                      Persona ON Producto_Proveedor.cPersona = Persona.cPersona INNER JOIN " & _
"                      Unidad ON Producto.cUnidad = Unidad.cUnidad ORDER BY DescripcionProducto"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptProductosProveedor", , App.Path + "\Reportes\")
End Sub

Private Sub mnuproductosproveedor1_Click()
FrmProductoProveedor.Show
End Sub

Private Sub mnuproductostransformaciones_Click()
FrmProductoTransformaciones.Show
End Sub

Private Sub MnuRegistroCompras_Click()
FrmReporteRegistroCompras.Show
End Sub

Private Sub mnurecibosmkl_Click()
frmVentasPagos.Show
End Sub

Private Sub mnuproyectos_Click()
frmmisproyectos.Show
End Sub

Private Sub mnuRegistroCompras1_Click()
FrmRegistroCompras.Show
End Sub

Private Sub mnuRegistroventas_Click()
FrmReporteRegistroVentas.Show
End Sub

Private Sub MnuReporteDetalleDinero_Click()
FrmReporteSalidaDinero.Show
End Sub

Private Sub mnuRegistroVentas1_Click()
FrmRegistroVentas.Show
End Sub

Private Sub MnuReporteKardex_Click()
FrmReporteKardex.Show
End Sub

Private Sub mnureportegeneral_Click()
frmReportesGenerales.Show
End Sub

Private Sub MnuRotacion_Click()
FrmRotacionProductos.Show
End Sub

Private Sub MnuRptClientes_Click()
FrmReportePersonas.Show
End Sub

Private Sub MnuRptCreditoDistribuidora_Click()
  FrmReporteCreditos.EnumRptCredito = CreditoDistribuidora
  FrmReporteCreditos.Show
End Sub

Private Sub MnuRptCreditosCliente_Click()
  
  FrmReporteProducto.Show
End Sub

Private Sub MnuRptDistribuidoras_Click()
Dim Ans As Boolean
  strCadena = "SELECT cDistribuidora, sRazonSocial,sDireccionDistribuidora, " & _
  " sRuc,sTelefonoDistribuidora1, sEmailDistribuidora FROM Distribuidora"
  Call ConfiguraRst(strCadena)
  Ans = ShowMultiReport(rst, "RptDistribuidora", , App.Path + "\Reportes\")
  Set rst = Nothing
End Sub

Private Sub MnuRptGanancias_Click()
  FrmReporteGanancia.Show
End Sub

Private Sub MnuRptKardex_Click()
  FrmReporteKardex.Show
End Sub

Private Sub MnuRptLaboratorio_Click()
Dim Ans As Boolean
  strCadena = "SELECT cLaboratorio, sRazonSocial,sDireccionLaboratorio1, " & _
  " sTelefonoLaboratorio1, sEmaillaboratorio FROM Laboratorio "
  Call ConfiguraRst(strCadena)
  Set DataReport3.DataSource = rst
    DataReport3.Show
  Set rst = Nothing
End Sub

Private Sub MnuRptLaboratorioDistribuidora_Click()
Dim Ans As Boolean
  strCadena = "SELECT Laboratorio.sRazonSocial as Laboratorio, Distribuidora.sRazonSocial " & _
  " As Distribuidora,lExclusivo FROM Laboratorio INNER JOIN (DistribuidoraLaboratorio " & _
  " INNER JOIN Distribuidora ON Distribuidora.cDistribuidora= " & _
  " DistribuidoraLaboratorio.cDistribuidora) ON DistribuidoraLaboratorio.cLaboratorio= " & _
  " Laboratorio.cLaboratorio"
  Call ConfiguraRst(strCadena)
  Ans = ShowMultiReport(rst, "RptLaboratorioDistribuidora", , App.Path + "\Reportes\")
  Set rst = Nothing
End Sub

Private Sub MnuRptProductos_Click()
  FrmReporteProducto.Show
End Sub

Private Sub MnuRptProductosComprar_Click()
  FrmReporteProductoCompra.Show
End Sub

Private Sub mnurregistrohuelladigital_Click()
FrmCapturaHuella.Show
Exit Sub
End Sub

Private Sub MnuSalidadeDinero_Click()
frmPrestamos.Show
End Sub

Private Sub MnuSalidaProductos_Click()
FrmReporteSalidas.Show
End Sub

Private Sub mnusalidaDinero_Click()
strCadena = "SELECT M.id_solicitud,M.fecha,U.nombre_usu,M.documentos,M.monto,M.detalle,MA.descripcion,D.cantidad,D.id_precio,D.total FROM materiales_solicitud M,usuario U,materiales_det D,materiales MA WHERE D.id_mat=MA.id_mat AND D.id_proyecto='00001' AND D.id_solicitud=M.id_solicitud AND   M.id_solicitud=D.id_solicitud AND D.id_creador='01185' AND  M.id_proyecto='00001' AND M.id_encargado=U.id_usu AND U.id_creador='01185'"
Call ConfiguraRstM(strCadena)
Ans = ShowMultiReport(rstM, "RptProyectos", , App.Path + "\Reportes\")
End Sub

Private Sub mnuSalir_Click()
  Call salir
  End
End Sub

Private Sub MnuSyser_Click()

End Sub

Private Sub MnuTipoDocumentos_Click()
FrmTipoDocumentos.Show
End Sub



Private Sub MnuSeguridad_Click()
frmDemo.Show
'FrmContactenos.Show
End Sub

Private Sub mnusolicitudespersonal_Click()

strCadena = "SELECT S.id_usu,U.nombre_usu,S.fecha_solicitud,monto_solicitado,monto_declarado,D.descripcion FROM solicitud_dinero S,usuario U,detalle_solicitud_dinero D WHERE S.num_solicitud=D.num_solicitud AND D.id_usu=S.id_usu AND S.id_usu=U.id_usu AND U.id_creador='01185' AND S.id_creador='01185' AND fecha_solicitud>='2013-01-01'"
Call ConfiguraRstM(strCadena)
Ans = ShowMultiReport(rstM, "RptSolicitudes", , App.Path + "\Reportes\")

End Sub

Private Sub mnuSolicitudViaticos_Click()
FrmSolicitudViaticos.Show
End Sub

Private Sub mnusublineas_Click()
FrmSublineas.Show
End Sub

Private Sub mnutanques_Click()
frmsurtidores.Show
End Sub

Private Sub mnuTipoCambio_Click()
FrmTipocambio.Show
End Sub

Private Sub MnuTipoMovimiento_Click()
  FrmTipoMovimiento.Show
End Sub

Private Sub MnuTipoProducto_Click()
 frmTipoProductoi.Show
End Sub

Private Sub MnuTrabajodores_Click()

End Sub

Private Sub mnutracking_Click()
frmtracking.Show
End Sub

Private Sub mnutransferencias10_Click()
FrmTransferencias.Show
End Sub

Private Sub MnuTransferenciasSunat_Click()
FrmTransferencias.Show
End Sub

Private Sub MnuTransportistas_Click()

End Sub

Private Sub mnuturnos00_Click()
FrmTurnos.Show
End Sub

Private Sub MnuUnidad_Click()
  FrmUnidad.Show
End Sub

Private Sub mnuUnidadesTransporte_Click()
FrmUnidadesTransporte.Show
End Sub

Private Sub MnuUsuarios_Click()
FrmUsuarios.Show
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case "(Producto)"
        FrmBuscarProducto.Show
    Case "(Compra)"
      EnumBuscar = BDocumentoCompra
      
    Case "(Venta)"
      EnumBuscar = BDocumentoVenta
    
        
  End Select
End Sub

Private Sub MnuUtilidadesDiarios_Click()
FrmEstadistica.Show
End Sub

Private Sub MnuValorizadoProductos_Click()
FrmReporteValorizado.Show
End Sub

Private Sub mnuvaraicion_Click()
frmVariacionCostes.Show
End Sub

Private Sub MnuVentadeProductos_Click()

FrmReporteRegistroVentas.Show

End Sub

Private Sub MnuZonas_Click()
FrmZonas.Show
End Sub

   
Private Sub StatusBar1_PanelClick(ByVal panel As MSComctlLib.panel)
Select Case panel.Index
        Case 1
            frmmenu.Show
             
            Exit Sub
        Case 3
            FrmDocumentos.Show
                         Exit Sub
        Case 4
            frmusuarioslinea.Show
        
        Case 6
            
            If KEY_ALARMA_STOCK = "si" Then
                strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,stock_minimo,precio_venta FROM view_producto WHERE stock<=stock_minimo and ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY id_linea,nombre_prod"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptStockMinimo", , App.Path + "\Reportes\")
            End If
            
            
            Exit Sub

        Case 7
            frmmistareas.Show
        
         Case 8
            If KEY_USUARIO = "42546269" Then
                Procedencia = seleccionar_per
                frmsegurity.Show
            End If
    End Select
End Sub


Public Function IsFormLoaded(fForm As Form, pDni As String) As Boolean


Dim X As Integer

For X = 0 To Forms.Count - 1

If (Forms(X) Is fForm) Then
  Dim frmaux As Frmchat
  Set frmaux = Forms(X)
  
  
  If frmaux.txtdni.Text = pDni Then
  
    IsFormLoaded = True

    frmaux.cerrar (rstChat("recibe"))
    'Frmchat.txtDNI.Text = rstChat("recibe")
    frmaux.txtdni.Text = rstChat("envia")
    frmaux.Caption = UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rstChat("envia")))
        
    frmaux.Show
          
    frmaux.llenar
    strCadena = "UPDATE chat set recd = 1 where id = '" & rstChat("id") & "'"
    CnBd.Execute strCadena
                
    
    
    
    Exit Function
  End If
  
End If

Next X

 IsFormLoaded = False

'Exit_Proc:
'Exit Function

'Err_Proc:
'MsgBox Err.Description
'Resume Exit_Proc


End Function

Private Sub timer_cloud_Timer()


End Sub





Private Sub Timer1_Timer()
Call buscar_solicitudes
End Sub
Public Sub buscar_solicitudes()
On Error GoTo salir
If KEY_USUARIO = "10002108" Then
    strCadena = "SELECT * FROM solicitud_credito WHERE ruc='" & KEY_RUC & "' and estado='01'"
    Call ConfiguraRstI(strCadena)
    If rstI.RecordCount > 0 Then
        PlaySound App.Path & "\sonidos\dingding.wav"
        FrmSolicitudCredito.Show
    End If
End If

Exit Sub
salir:

End Sub

Private Sub Tiner_update_Timer()
Call verificar_versionsoft
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case (Button.Index)
    Case 1
        If KEY_RUBRO = "00025" Then
            FrmMatricula.Show
        Else
            FrmPersona.Show
        End If
    Case 3
        frmpersonal.Show
    Case 5
       FrmVentas.Show
    Case 7
         FrmCompras.Show
    Case 9
         FrmTransferencias.Show
    Case 7
        FrmTransferencias.Show
   Case 9
        FrmMiscuentas.Show
    Case 11
        If KEY_RUBRO = "00027" Then
            frmLibro.Show
            Exit Sub
        Else
        FrmProducto.Show
        End If
   Case 12
        Dim ejecuta As Double
      '  ejecuta = Shell("I:\Archivos de programa\3CXPhone\3CXPhone.exe", vbNormalFocus)
    Case 13
        FrmReporteRegistroVentas.Show
        Exit Sub
    Case 15
        frmpersonaasistencia.Show
        Exit Sub
    
    Case 19
        frmMemorandun.Show
        Exit Sub
    Case 21
        'FrmHuellaDigital.Show
        frmCorProcesos.Show
        Exit Sub
     Case 23
        FrmImpTramite.Show
        'frmCorTramite.Show
        Exit Sub
        
    Case 25
        FrmServiciotecnico.Show
        'frmmantenimientos.Show
        Exit Sub
    Case 29
        FrmAdmisionEmergencia.Show
        Exit Sub
    Case 31
        
        Call salir
   
         
End Select
End Sub

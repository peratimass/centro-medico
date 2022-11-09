Attribute VB_Name = "ModVariables"

Option Explicit
Public CnBd As ADODB.Connection
Public CnBd2 As ADODB.Connection
Public CnBd3 As ADODB.Connection
Public rst As ADODB.Recordset
Public rst2 As ADODB.Recordset
Public rst3 As ADODB.Recordset
Public rstA As ADODB.Recordset
Public rstChat As ADODB.Recordset
Public rstT As ADODB.Recordset
Public rstK As ADODB.Recordset
Public rstI As ADODB.Recordset
Public rstM As ADODB.Recordset
Public rstL As ADODB.Recordset
Public rstZ As ADODB.Recordset
Public rstF As ADODB.Recordset
Public rstPP As ADODB.Recordset
Public rstc As ADODB.Recordset
Public rstIN As ADODB.Recordset
Public rstUpdate As ADODB.Recordset
Public rstAlarma As ADODB.Recordset
Public rstVersion As ADODB.Recordset



Public rstAux As ADODB.Recordset
Public rstchat2 As ADODB.Recordset
Public RstEjecuta As ADODB.Recordset
Public rstTemporal As ADODB.Recordset
Public rstCloud As ADODB.Recordset
Public rstLocal As ADODB.Recordset
Public rstP As ADODB.Recordset
Public strRuta As String
Public strCadena As String
Public strSQL As String

Public HfdPrecio As MSHFlexGrid
Public DteFecha As Date
Public longitud As Integer
Public KeyAsc As Integer
Public IntCantMov As Integer
Public StrRpta As String
Public strtabla As String
Public strTipo As String
Public StrCodProducto As String
Public StrTipoMov As String
Public Lnumdocu As String
Public Gencodigo As String
Public StrNumero As String
Public strServer_ini As String
Public strRuta_ini As String
Public strRuta_reloj As String
Public KEY_CONTABILIDAD As String
Public KEY_PASSWORD As String
Public KEY_DETALLE As String
Public KEY_MONTOENVASE As Single
Public KEY_ENVASE As String
Public KEY_DIR_PUBLIC As String
Public KEY_CHECK_STOCK_FACTURA As String
Public cNaturaleza As String
Public ccostos1G As String
Public ccostos2G As String
Public ccostos3G As String
Public ccostos4G As String
Public KEY_TELEFONO As String
Public KEY_PROYECTO As String
Public Const KEY_NULO As Double = 0
Public KEY_VENDEDOR As String
Public KEY_IGV As Single
Public KEY_EMPRESA As String
Public KEY_AUTOMATICO As String
Public KEY_CERVECERIA As String
Public KEY_UPDATE_PRECIOS As String
Public KEY_RUC As String
Public KEY_DIRECCION As String
Public KEY_DIRECCION_ALM As String
Public KEY_BARRAS As String
Public KEY_COMPROBANTE As String
Public KEY_FOTO As String
Public KEY_TRAMITE As String
Public KEY_CAJA_INDEPENDIENTE As String
Public KEY_SUCURSAL As String
Public KEY_VENTANILLA As String
Public KEY_FACTURACION_DETALLADA As String
Public KEY_FACTURACION_CENTRALIZADA As String
Public KEY_LINEA_CREDITO  As String
Public KEY_VERSION As String
Public KEY_GRIFO As String
Public KEY_SKIN As String
Public KEY_COLOR_BARRA As String
Public KEY_GUIA_FRACCIONADA As String
Public KEY_UPDATE_PROFORM As String
Public KEY_PAIS As String
Public KEY_BANDERA  As String
Public KEY_MONEDA As String
Public KEY_RUC_VERSION As String

Public KEY_ALARMA_STOCK As String
Public KEY_RESERVA_STOCK  As String

Public KEY_DESCUENTO_LINEA As String

'****** IMPUESTO BOLSAS ****

Public KEY_IMPUESTO_BOLSAS As String
Public KEY_VALOR_BOLSA As Single

Public KEY_CAMBIO_PRECIO_PASS  As String
Public KEY_SEGMENTACION_PRECIO  As String

Public KEY_COMPROBANTES_PROPIOS As String

Public KEY_COMPROBANTE_ADICIONAL As String
Public KEY_VENDEDOR_AMBULANTE As String
Public KEY_STOCK_GLOBAL As String

Public KEY_CTA_COMPRA_SOLES As String
Public KEY_CTA_COMPRA_DOLARES As String
Public KEY_CTA_COMPRA_RH  As String


'Cuentas Nuevas
Public KEY_CTA_PAGAR_SERVICIO  As String
Public KEY_CTA_IGV_VENTA As String
Public KEY_CTA_IGV_SERVICIO_COMPRA As String
Public KEY_BONIFICACIONES As String
Public KEY_EMPRESA_PLAN  As String
Public KEY_SIN_EFECTO_CAJA As String
Public KEY_FECHA_CORTE As Date

'letra por pagar
Public KEY_CTA_LETRA_PAGAR_SOLES  As String
Public KEY_CTA_LETRA_PAGAR_DOLARES  As String

'Intrumentos de descuento FET
Public KEY_CTA_FET_SOLES  As String
Public KEY_CTA_FET_DOLARES  As String

Public KEY_MOVIMIENTO_SIN_STOCK As String


'Cuentas pagar Anticipos
Public KEY_CTA_ANT_SOLES  As String
Public KEY_CTA_ANT_DOLARES  As String


Public KEY_TURNO As String
Public KEY_ALERTA_CORTE As String
Public KEY_GRUPO_EMPRESARIAL As String


Public KEY_MOSTRAR_PRECIO_COSTO  As String
Public KEY_MOSTRAR_PRECIO_MAYOR As String

Public KEY_MOSTRAR_SUCURSAL As String


Public Const KEY_CANCELADO As String = "NN"
Public Const KEY_PENDIENTE As String = "PP"
Public Const KEY_CANCELAR As String = "(Cancelar)"
Public Const KEY_ANULV As String = "V"
Public Const KEY_ANULF As String = "F"
Public Const KEY_MONTO As String = "(Monto)"
Public Const KEY_SUPERVISOR As String = "0001"
Public Const KEY_AUTOR As String = "0007"
Public Const KEY_PERU As String = "9589"

Public KEY_PORCENTAJE_ZONA As Single
Public KEY_PORCENTAJE_CREDITO As Single


Public KEY_HABILITADO_NOTACREDITO As String
Public KEY_IMPRESION_PROFORMA As String




Public KEY_DEPARTAMENTO As String
Public KEY_PROVINCIA As String
Public KEY_DISTRITO  As String
Public KEY_DETALLE_COMBO As String

Public KEY_RUBRO As String
Public KEY_FECHA  As String
Public KEY_CARGO As String

Public KEY_CAMBIO As Single

Public KEY_CAMBIO_VENTA As Single
Public KEY_CAMBIO_COMPRA As Single
Public KEY_CAMBIO_LOCAL As Single
Public KEY_CONVERSION_CAMBIO As String

Public KEY_STOCK_CONTABLE As String
Public KEY_ASIENTO_GLOBAL_CTA_PAGAR As String

Public KEY_TIPO_LETRA As String

Public KEY_CAMARA As String
Public KEY_IMPRESORA As String
Public KEY_PRINTER As String
Public KEY_PRINTER_GUIA As String

Public KEY_CADENA As String
Public KEY_CLOUD As String

Public KEY_TOKEN_CLOUD  As String
Public KEY_SERVIDOR_KEYFACIL As String
Public KEY_TOKEN_LOCAL  As String
Public KEY_TOKEN_SUCURSAL As String
Public KEY_CODIGO_UNIVERSAL_IMPRESION  As String
Public KEY_NOMBRE_COMERCIAL As String




Public KEY_VALIDACION_EXTREMA  As String
Public KEY_PAQUETE_EMPRESARIAL  As String
Public KEY_GENERADOR_MENSUALIDAD As String
Public KEY_CONTROL_MERCADERIA  As String

Public KEY_CTA_COBRAR_PRODUCTO  As String
Public KEY_CTA_COBRAR_SERVICIO  As String

Public KEY_CTA_INGRESO_PRODUCTO  As String
Public KEY_CTA_INGRESO_SERVICIO  As String

Public KEY_PRODUCTO_REPETIDO As String
Public KEY_DIAS_CREDITO As Integer
Public KEY_MORA As String
Public KEY_MORA_MONTO As Single
Public KEY_PROVEEDOR As String
Public KEY_SERVIDOR_CLOUD As String

Public KEY_NOTA_CREDITO_ADMIN As String
Public KEY_NOTA_CREDITO_USER As String

Public KEY_REFERENCIA_COMPROBANTE As String


Public Const KEY_IMPORTAR As String = "(Importar)"
Public Const KEY_EXPORTAR As String = "(Exportar)"
Public KEY_SKFACTURA As String
Public KEY_CON_IGV As String
Public KEY_APLICA_IGV As String * 2
Public KEY_FINGERPRINT As String * 2
Public KEY_SERIE_DEFAULT As String
Public Const KEY_TECNICO As String = "(Tecnico)"
Public Const KEY_ADMINISTRADOR As String = "00004"
Public Const KEY_CLIENTE As String = "00000000"
Public Const KEY_ADELANTO As String = "0003"
Public Const KEY_CREDITO As String = "0004"
Public Const KEY_CONTADO As String = "0001"
Public Const KEY_SALDINER As String = "0097"
Public Const KEY_INGDINER As String = "0016"
Public Const KEY_NOTACRED As String = "0098"
Public Const KEY_NOTAPED As String = "0088"
Public Const KEY_COTIZA As String = "0109"
Public Const KEY_GRINTER As String = "0095"
Public Const KEY_PAGOCRE As String = "0010"
Public Const KEY_NEW As String = "(Nuevo)"
Public Const KEY_MAIL As String = "(Mail)"
Public Const KEY_UPDATE As String = "(Modificar)"
Public Const KEY_SECTOR As String = "(Sector)"
Public Const KEY_PENDIENT As String = "(Pendiente)"
Public Const KEY_DELETE As String = "(Eliminar)"
Public Const KEY_BROWSER As String = "(Buscar)"
Public Const KEY_SAVE As String = "(Grabar)"
Public Const KEY_CANCEL As String = "(Cancelar)"
Public Const KEY_URBANIZACION As String = "(Urbanizacion)"
Public Const KEY_ZONA As String = "(Zona)"
Public Const KEY_EXIT As String = "(Salir)"
Public Const KEY_ORDENPAGO As String = "(OrdenPago)"
Public Const KEY_REVERTIR As String = "(Revertir)"
Public Const KEY_MERMAS As String = "(Mermas)"
Public Const KEY_PRINT As String = "(Imprimir)"
Public Const KEY_AGREGAR As String = "(Agregar)"
Public Const KEY_ASIGNAR As String = "(Asignar)"
Public Const KEY_QUITAR As String = "(Quitar)"
Public Const KEY_OK As String = "(Aceptar)"
Public Const KEY_EXCEL As String = "(Exportar)"
Public Const KEY_GUIAREMISION As String = "(GuiaRemision)"
Public Const KEY_ANULAR As String = "(Anular)"
Public Const KEY_ACTUALIZAR As String = "(Actualizar)"
Public Const KEY_DEUDAS As String = "(Deudas)"
Public Const KEY_PAGAR As String = "(Pagar)"
Public Const KEY_ATENDER As String = "(Atender)"
Public Const KEY_REPORTE As String = "(Reporte)"
Public Const KEY_RVENTAS As String = "(RVentas)"
Public Const KEY_RCOMPRAS As String = "(RCompras)"
Public Const KEY_RGASTOS As String = "(Gastos)"
Public Const KEY_SAL As String = "S01"
Public Const KEY_SALNULL As String = "S02"
Public Const KEY_ING As String = "I01"
Public Const KEY_TRANS As String = "T01"
Public Const KEY_TRANSNULL As String = "T02"
Public Const KEY_INGNULL As String = "I02"
Public Const KEY_PERNAT As String = "N"
Public Const KEY_PERJU As String = "J"
Public Const KEY_HUELLA As String = "(key_huella)"
Public Const KEY_COD_PER As String = "02484"
Public Const KEY_PRODUCTO_INTERES  As String = "06225"

Public KEY_INVETARIO_FACTURA As String
Public KEY_USUARIO As String
Public KEY_EMAIL As String
Public KEY_SEGURO_VENTA As String
Public KEY_ENVIO_SUNARP  As String
Public KEY_ALM As String

Public KEY_TRANSPORTE_MIGRA As String


Public KEY_PORCENTAJE_INTERES As Single




Public Const KEY_GUIA As String = "0009"
Public Const KEY_FACTURA As String = "0001"
Public Const KEY_INGALMA As String = "0089"
Public Const KEY_BOLETA As String = "0003"
Public Const KEY_PEDIDO As String = "0103"
Public Const KEY_RBOINGRESO As String = "0108"
Public Const KEY_RBOEGRESO As String = "0097"
Public Const KEY_DSCTO As Single = 0#

Public Const KEY_MAQUINA1 As String = "FFFF101307"
Public KEY_MODELO_COLOR  As String
Public KEY_FACTURACION_ELECTRONICA As String
Public KEY_RESOLUCION As String
Public KEY_TRACKING  As String
Public KEY_AGRANEL As String

Public Const KEY_SERIE1 As String = "001"
Public Const KEY_SERIE2 As String = "002"
Public Const KEY_SERIE3 As String = "003"
Public Const KEY_SERIE4 As String = "004"
Public Const KEY_SERIE5 As String = "005"
Public Const KEY_SERIE6 As String = "006"
Public Const KEY_SERIE7 As String = "007"

Public KEY_CTA_DETRACCION As String
Public KEY_PORCENTAJE_DETRACCION As Single

Public Const KEY_ADMIN As String = "00004"
Public Const KEY_SUPER As String = "00009"
Public Const KEY_CAJA As String = "00008"
Public key_cajas As Integer
Public Const DTEMAXIMA As Date = #12/31/2050#
Public Const DTEMINIMA As Date = #1/1/1900#
Public Const DTEDEFECTO As Date = #1/1/1900#

Public Const MSGVALIDACION As String = "Validación del Sistema"
Public Const MSGVERIFICACION As String = "Verificación del Sistema"
Public Const MSGERROR As String = "Error del Sistema"
Public Const MSGGRABACION As String = "Grabación no realizada"
Public Const MSGVENCIMIENTO As String = "Fecha de Vencimiento"

Public Const MSGELIMINAR As String = "¿Desea Eliminar el Registro?"
Public Const MSGANULAR As String = "¿Desea Anular el Documento?"
Public Const MSGCANCELAR As String = "¿Desea Cancelar todo el Movimiento?"
Public Const MSGCONGUIA As String = "¿Tiene la Factura de los Producto?"
Public Const MSGSTAND As String = "Ese Producto ya está en Stand."
Public Const MSGAGREGAR As String = "¿Desea agregarlo?"
Public Const MSGBAJAPRODUCTOS As String = "Hay productos vencidos, faltantes o están por terminarse."
Public Const MSGVERPRODUCTOS As String = "¿Desea ver los productos?"
Public Const MSGNOTCLIENTE As String = "La persona seleccionada no es un cliente como para darle un crédito."
Public Const MSGENTIDAD As String = "Debe seleccionar una entidad."
Public Const MSGFALTAPRODUCTOS As String = "No hay suficientes existencias del producto."
Public Const MSGFALTADATOS As String = "Complete los Datos."
Public Const MSGDATOS As String = "Complete Bien los Datos."
Public Const MSGDUPLICIDADFACTURA As String = "Presenta Duplicidad en el Nº de Factura."
Public Const MSGDUPLICIDAD As String = "Presenta Duplicidad de Productos."
Public Const MSGLOTE As String = "Actualizar los Lotes de Productos."
Public Const MSGTAMAÑO As String = "El número tiene una longitud superior de la permitida."
Public Const MSGASIENTO As String = "Ya se realizó el asiento contable respectivo."
Public Const MSGREINGRESEDATOS As String = "Ingrese nuevamente los registros."
Public Const MSGFECHA As String = "Ingrese la fecha de Vencimiento del Documento:"
Public Const MSGNOTANULAR As String = "No se puede Anular el Documento"
Public Const MSGNOTPAGAR As String = "No es posible, el Documento ya fue pagado."
Public Const MSGNOTATENDER As String = "No es posible, el Pedido ya fue atendido o fue anulado."
Public Const MSGSTOCK As String = "NO CUENTA CON STOCK NECESARIO PARA REALIZAR LA VENTA"

Public Const KEY_RUTA_SERVER As String = "repos.vitekey.com/tools/scrum/"
Public vx_sincont As String

'***** VARIABLES LECTOR DE HUELLAS
Option Base 0
Public Const MAX_USERID_SIZE As Long = 50
Public Const MAX_TEMPLATE_SIZE As Long = 1024
Public Const MAX_MEMO_SIZE As Long = 100
Public Const DATABASE_COL_SERIAL As Long = 0
Public Const DATABASE_COL_USERID As Long = 1
Public Const DATABASE_COL_FINGERINDEX As Long = 2
Public Const DATABASE_COL_TEMPLATE1 As Long = 3
Public Const DATABASE_COL_TEMPLATE2 As Long = 4
Public Const DATABASE_COL_MEMO As Long = 5
Public m_hScanner As Long
Public m_hDatabase As Long
Public m_hMatcher As Long
Public m_strError As String
Public m_Serial As Long
Public m_UserID As String
Public m_FingerIndex As Long
Public m_Template1(MAX_TEMPLATE_SIZE - 1) As Byte
Public m_Template1Size As Long
Public m_Template2(MAX_TEMPLATE_SIZE - 1) As Byte
Public m_Template2Size As Long
Public m_Memo As String
'*********************************





Public Enum EnumProcede
  nuevo = 1
  modificar
  Eliminar
  Selecionar
  anular
  Neutro
  relacionar
  combo
  mermas
  buscar
  pendientes
  transformaciones
  revertir
  imprimir_s
  seleccionar_soldadura
  seleccionar_ensamblaje
  seleccionar_tapiz
  seleccionar_otro
  seleccionar_vendedor
  seleccionar_per
  seleccionar_insumo
  eliminar_insumo
  modificar_credito
  seleccionar_atencion
  eliminar_informe
  modificar_precio
  mailenviar
  registroventadetalle
  modificar_precio_unitario
  cerrarcaja
  prorrateo_importacion
  prorrateo_gasto
  extornar
  diferida
End Enum

Public Enum EnumCostos
  Buscar1 = 1
  Buscar2
  Buscar3
  Buscar4
  Buscar5
  Buscar6
  NeutroCostos
  End Enum

Public Enum EnumFactura
    Nueva = 1
    Modifica
End Enum
Public Enum EnumGuia
  NuevaGuia = 1
  ModificarGuia
  EliminarGuia
  MostrarGuia
End Enum
Public Enum EnumCliente
  buscarcliente = 1
  DetalleCliente
  DocumentoVenta
  Pedido
  COtraEntrada
  COtraSalida
End Enum
Public Enum EnumProducto
  BuscarProducto = 1
  DetalleProducto
  PDocumentoCompra
End Enum

'Public EnumAlmacen
  'NuevoAlmacen
  
  'DocumentoCompra
  'DOtraEntrada
  'DOtraSalida
'End Enum

Public Enum EnumBuscarDocumento
  BPedido = 1
  BDocumentoVenta
  BDocumentoCompra
  BOtraEntrada
  BOtraSalida
End Enum
Public Enum prntAlineacion
    pAlnIzquierda = 1
    pAlnDerecha = 2
    pAlnCentro = 3
End Enum

Public Enum EnumCredito
  CreditoCliente = 1
  CreditoDistribuidora
End Enum

Public Enum debe_haber
    debe = 1
    HABER
End Enum
Public Enum Descuento
    unitario = 1
    Total
    vacio
End Enum
Public Enum Calculadora
        Mbimponible = 1
        MbimponibleInafecta
        Migv
        MOtro
        Misc
        MTC
        Mneutro
End Enum


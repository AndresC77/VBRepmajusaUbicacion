VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NEED - ED"
   ClientHeight    =   4725
   ClientLeft      =   -165
   ClientTop       =   1050
   ClientWidth     =   9825
   Icon            =   "mdiPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdiPrincipal.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4350
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "Empresa - Sucural"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "Punto de Facturación"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3969
            Object.ToolTipText     =   "ENLACE DIGITAL"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "Usuario del sistema"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:55"
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2018-11-08"
            Object.ToolTipText     =   "Fecha"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu menEmpresa 
         Caption         =   "&Empresa"
         HelpContextID   =   1
      End
      Begin VB.Menu menNuevaEmpresa 
         Caption         =   "&Nueva Empresa"
         HelpContextID   =   1
      End
      Begin VB.Menu menModEmpresa 
         Caption         =   "&Modificar Empresa"
         HelpContextID   =   1
      End
      Begin VB.Menu menNuevoBrowser 
         Caption         =   "Nuevo Browser"
         HelpContextID   =   1
      End
      Begin VB.Menu menS1_1 
         Caption         =   "-"
      End
      Begin VB.Menu menCambioClave 
         Caption         =   "&Cambiar Clave"
         HelpContextID   =   1
      End
      Begin VB.Menu menSalir 
         Caption         =   "&Salir"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menVentana 
      Caption         =   "Ve&ntana"
      WindowList      =   -1  'True
      Begin VB.Menu menMoshor 
         Caption         =   "Mosaico &Horizontal"
         HelpContextID   =   1
      End
      Begin VB.Menu menMosver 
         Caption         =   "Mosaico &Vertical"
         HelpContextID   =   1
      End
      Begin VB.Menu menCascada 
         Caption         =   "&Cascada"
         HelpContextID   =   1
      End
      Begin VB.Menu menOrgico 
         Caption         =   "&Organizar Iconos"
         HelpContextID   =   1
      End
      Begin VB.Menu menS2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu menAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu menAcercade 
         Caption         =   "&Acerca de..."
         HelpContextID   =   1
      End
      Begin VB.Menu menS3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu menContabilidad 
      Caption         =   "&Contabilidad"
      Begin VB.Menu menImpCompConta 
         Caption         =   "&Ver Comprobantes"
         HelpContextID   =   1
      End
      Begin VB.Menu menContabilizarMovMer 
         Caption         =   "&Contabilizar Movimientos Mercadería"
         HelpContextID   =   1
      End
      Begin VB.Menu menS4 
         Caption         =   "-"
      End
      Begin VB.Menu menPlanCuenta 
         Caption         =   "&Plan de Cuentas"
         HelpContextID   =   1
      End
      Begin VB.Menu menCentroCosto 
         Caption         =   "Centros de Costo"
         HelpContextID   =   1
      End
      Begin VB.Menu menDefIVA 
         Caption         =   "&Definición de IVA"
         HelpContextID   =   1
         Begin VB.Menu menDefIVAC 
            Caption         =   "Para &Compras"
            HelpContextID   =   2
         End
         Begin VB.Menu menDefIVAV 
            Caption         =   "Para &Ventas"
            HelpContextID   =   2
         End
         Begin VB.Menu menS4_1 
            Caption         =   "-"
         End
      End
      Begin VB.Menu menCtaMov 
         Caption         =   "Cta. Contable de &Movimientos"
         HelpContextID   =   1
         Begin VB.Menu menCtaMovIng 
            Caption         =   "Documentos de &Ingreso"
            HelpContextID   =   2
         End
         Begin VB.Menu menCtaMovEgr 
            Caption         =   "Documentos de &Egreso"
            HelpContextID   =   2
         End
         Begin VB.Menu menS4_2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu menRetenciones 
         Caption         =   "&Retenciones"
         HelpContextID   =   1
      End
      Begin VB.Menu menDefFormaPago 
         Caption         =   "Definición de &Formas de Pago"
         HelpContextID   =   1
      End
      Begin VB.Menu menDefCXPXC 
         Caption         =   "Defi&nición de Cta. Conta"
         HelpContextID   =   1
         Begin VB.Menu menDefCXC 
            Caption         =   "Para Cta. por &Cobrar"
            HelpContextID   =   2
         End
         Begin VB.Menu menDefCXP 
            Caption         =   "Para Cta. por &Pagar"
            HelpContextID   =   2
         End
         Begin VB.Menu menS4_3 
            Caption         =   "-"
         End
      End
      Begin VB.Menu menSucursal 
         Caption         =   "Sucursales"
         HelpContextID   =   1
      End
      Begin VB.Menu menCambioFac 
         Caption         =   "Cambio Datos de Facturas"
         HelpContextID   =   1
      End
      Begin VB.Menu menAnularFac 
         Caption         =   "Anulación de Facturas"
         HelpContextID   =   1
      End
      Begin VB.Menu menAnularCom 
         Caption         =   "Anulación de Compras Locales"
         HelpContextID   =   1
      End
      Begin VB.Menu menImpAsiento 
         Caption         =   "Imprimir Asientos"
         HelpContextID   =   1
      End
      Begin VB.Menu menQuitarNotaCredito 
         Caption         =   "Anular/Desaplicar Notas Credito"
         HelpContextID   =   1
      End
      Begin VB.Menu menCierreMes 
         Caption         =   "Cierre de Mes"
         HelpContextID   =   1
      End
      Begin VB.Menu menReprocesoRet 
         Caption         =   "Reprocesos de Retencion"
         HelpContextID   =   1
      End
      Begin VB.Menu menCuadreAsientos 
         Caption         =   "Cuadre de Asientos"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menCliPro 
      Caption         =   "Clientes/&Proveedores"
      Begin VB.Menu menCatalogosCliProv 
         Caption         =   "Catalogos"
         HelpContextID   =   1
         Begin VB.Menu menCatCli 
            Caption         =   "&Categoria de Clientes"
            HelpContextID   =   2
         End
         Begin VB.Menu menCatPro 
            Caption         =   "C&ategoria de Proveedores"
            HelpContextID   =   2
         End
         Begin VB.Menu menUbicacion 
            Caption         =   "&Ubicacion"
            HelpContextID   =   2
            Begin VB.Menu menPais 
               Caption         =   "&País"
               HelpContextID   =   3
            End
            Begin VB.Menu menCanton 
               Caption         =   "C&anton"
               HelpContextID   =   3
            End
            Begin VB.Menu menCiudad 
               Caption         =   "&Ciudad"
               HelpContextID   =   3
            End
            Begin VB.Menu menS5_1 
               Caption         =   "-"
            End
            Begin VB.Menu menZona 
               Caption         =   "&Zona"
               HelpContextID   =   3
            End
            Begin VB.Menu menRegion 
               Caption         =   "Region"
               HelpContextID   =   3
            End
         End
         Begin VB.Menu menS5_2 
            Caption         =   "-"
         End
         Begin VB.Menu menLineaN 
            Caption         =   "&Línea de Negocio"
            HelpContextID   =   2
         End
         Begin VB.Menu menTipoNegcio 
            Caption         =   "Tipos de &Negocio"
            HelpContextID   =   2
         End
         Begin VB.Menu menTipDoc 
            Caption         =   "&Tipos de Documentos"
            HelpContextID   =   2
         End
         Begin VB.Menu menFormaEntrega 
            Caption         =   "&Formas de Entrega"
            HelpContextID   =   2
         End
         Begin VB.Menu mencanal 
            Caption         =   "Canales"
            HelpContextID   =   2
         End
         Begin VB.Menu menSac 
            Caption         =   "SAC's"
            HelpContextID   =   2
         End
         Begin VB.Menu menCobrador 
            Caption         =   "Cobradores"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS5 
         Caption         =   "-"
      End
      Begin VB.Menu menCliente 
         Caption         =   "Clientes"
         HelpContextID   =   1
      End
      Begin VB.Menu menClientesPVP 
         Caption         =   "Clientes PVP UIO"
         HelpContextID   =   1
      End
      Begin VB.Menu menClientesPVPREC 
         Caption         =   "Clientes PVP RECREO"
      End
      Begin VB.Menu menClientesPVPGYE 
         Caption         =   "Clientes PVP GYE"
         HelpContextID   =   1
      End
      Begin VB.Menu menActualizacionClientes 
         Caption         =   "Actualizacion de Datos Clientes"
         HelpContextID   =   1
      End
      Begin VB.Menu menClientesPromo 
         Caption         =   "Clientes Promo UIO"
         HelpContextID   =   1
      End
      Begin VB.Menu menClientesPromoGYE 
         Caption         =   "Clientes Promo GYE"
         HelpContextID   =   1
      End
      Begin VB.Menu menClienteMod 
         Caption         =   "Modificar Clientes"
         HelpContextID   =   1
      End
      Begin VB.Menu menClienteModJefeSAC 
         Caption         =   "Modificar Clientes (JefeSAC)"
         HelpContextID   =   1
      End
      Begin VB.Menu menClienteModSAC 
         Caption         =   "Modificar Clientes (SAC)"
         HelpContextID   =   1
      End
      Begin VB.Menu menClienteModCartera 
         Caption         =   "Modificar Clientes (Cartera)"
         HelpContextID   =   1
      End
      Begin VB.Menu menClienteModGerencia 
         Caption         =   "Modificar Clientes (Gerencia)"
         HelpContextID   =   1
      End
      Begin VB.Menu menProveedor 
         Caption         =   "Proveedores"
         HelpContextID   =   1
      End
      Begin VB.Menu menFidelizacion 
         Caption         =   "Fidelización"
         HelpContextID   =   1
         Begin VB.Menu menCargaClientesFidelizacion 
            Caption         =   "Cargar Clientes"
            HelpContextID   =   2
         End
         Begin VB.Menu menS5_5 
            Caption         =   "-"
         End
         Begin VB.Menu menEnvioCorreo 
            Caption         =   "Envio de Correo a Clientes"
            HelpContextID   =   2
         End
      End
   End
   Begin VB.Menu menInventario 
      Caption         =   "&Inventario"
      Begin VB.Menu menCatalogosInv 
         Caption         =   "&Catálogos"
         HelpContextID   =   1
         Begin VB.Menu menMotivoAjuste 
            Caption         =   "Moti&vo de Ajuste"
            HelpContextID   =   2
         End
         Begin VB.Menu menLineaP 
            Caption         =   "&Línea de Producto"
            HelpContextID   =   2
         End
         Begin VB.Menu menMarca 
            Caption         =   "&Marca"
            HelpContextID   =   2
         End
         Begin VB.Menu menVerGrupos 
            Caption         =   "&Grupos de Productos"
            HelpContextID   =   2
         End
         Begin VB.Menu menUnidad 
            Caption         =   "&Unidad de Medida"
            HelpContextID   =   2
         End
         Begin VB.Menu menTalla 
            Caption         =   "&Talla"
            HelpContextID   =   2
         End
         Begin VB.Menu menColor 
            Caption         =   "&Color"
            HelpContextID   =   2
         End
         Begin VB.Menu menColeccion 
            Caption         =   "Col&ección"
            HelpContextID   =   2
         End
         Begin VB.Menu menAdmProducto 
            Caption         =   "&Productos"
            HelpContextID   =   2
         End
         Begin VB.Menu menProductoClc 
            Caption         =   "Producto Colección"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_1 
            Caption         =   "-"
         End
         Begin VB.Menu menListaPrecio 
            Caption         =   "L&ista de Precio"
            HelpContextID   =   2
         End
         Begin VB.Menu menDeposito 
            Caption         =   "&Bodegas"
            HelpContextID   =   2
         End
         Begin VB.Menu menOtroCargo 
            Caption         =   "&Otros Recargos"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_2 
            Caption         =   "-"
         End
         Begin VB.Menu menUbicacionBodega 
            Caption         =   "Ubicaciones en bodega"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS6 
         Caption         =   "-"
      End
      Begin VB.Menu menMantGuiaProveedor 
         Caption         =   "Mantener Guias de Proveedor"
         HelpContextID   =   1
      End
      Begin VB.Menu menS6_3 
         Caption         =   "-"
      End
      Begin VB.Menu menListaProducto 
         Caption         =   "Lista de Productos"
         HelpContextID   =   1
      End
      Begin VB.Menu menVerProductos 
         Caption         =   "Ver Productos"
         HelpContextID   =   1
      End
      Begin VB.Menu menAdmMovimiento 
         Caption         =   "Administrar Movimientos"
         HelpContextID   =   1
      End
      Begin VB.Menu menVerMovimiento 
         Caption         =   "Ver Movimientos"
         HelpContextID   =   1
      End
      Begin VB.Menu menIngresos 
         Caption         =   "Ingresos"
         HelpContextID   =   1
         Begin VB.Menu menComprasLocales 
            Caption         =   "Compras Locales"
            HelpContextID   =   2
         End
         Begin VB.Menu menNotasCreditoCliente 
            Caption         =   "Notas de Crédito Cliente"
            HelpContextID   =   2
         End
         Begin VB.Menu manGuiasProveedor 
            Caption         =   "Guias de Proveedor"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_4 
            Caption         =   "-"
         End
         Begin VB.Menu menAltasAuditoria 
            Caption         =   "Altas de Auditoría"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menEgresos 
         Caption         =   "Egresos"
         HelpContextID   =   1
         Begin VB.Menu menNotasCreditoProveedor 
            Caption         =   "Notas de Crédito Proveedor"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_5 
            Caption         =   "-"
         End
         Begin VB.Menu menBajasAuditoria 
            Caption         =   "Bajas de Auditoria"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menTransformacion 
         Caption         =   "Transformaciones"
         HelpContextID   =   1
      End
      Begin VB.Menu menTransfBodega 
         Caption         =   "Transferencia a Bodega"
         HelpContextID   =   1
      End
      Begin VB.Menu menCambios 
         Caption         =   "Cambios"
         HelpContextID   =   1
         Begin VB.Menu menCambioProductoTotal 
            Caption         =   "Ingreso Cambios de Productos TOTAL"
            HelpContextID   =   2
         End
         Begin VB.Menu menCambioProducto 
            Caption         =   "Ingreso Cambios de Productos"
            HelpContextID   =   2
         End
         Begin VB.Menu menCambioProductoPVPUIO 
            Caption         =   "Ingreso Cambios de Productos PVP UIO"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_8 
            Caption         =   "-"
         End
         Begin VB.Menu menRealizaCambioProducto 
            Caption         =   "Realizar Cambios de Producto"
            HelpContextID   =   2
         End
         Begin VB.Menu menRealizaCambioProductoPVPUIO 
            Caption         =   "Realizar Cambios de Producto PVP UIO"
            HelpContextID   =   2
         End
         Begin VB.Menu menRealizaDevolucionProducto 
            Caption         =   "Realizar Devoluciones de Producto"
            HelpContextID   =   2
         End
         Begin VB.Menu menRealizaDevolucionProductoPVPUIO 
            Caption         =   "Realizar Devoluciones de Producto PVP UIO"
            HelpContextID   =   2
         End
         Begin VB.Menu menDesmantelar 
            Caption         =   "Desmantelar"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS6_6 
         Caption         =   "-"
      End
      Begin VB.Menu menInventarioFisico 
         Caption         =   "Inventario Físico"
         HelpContextID   =   1
         Begin VB.Menu menVerIngInventario 
            Caption         =   "Ver Conteos de Inventario"
            HelpContextID   =   2
         End
         Begin VB.Menu menS6_7 
            Caption         =   "-"
         End
         Begin VB.Menu menAjusteInventario 
            Caption         =   "Ajuste de Inventario"
            HelpContextID   =   2
         End
         Begin VB.Menu memAdminContenedores 
            Caption         =   "Adm.Contenedor"
            HelpContextID   =   2
         End
         Begin VB.Menu memIngContenedores 
            Caption         =   "Ingresar Contenedores"
            HelpContextID   =   2
         End
         Begin VB.Menu memIngContenedoresDesc 
            Caption         =   "Ingresar Contenedores DESC"
            HelpContextID   =   2
         End
         Begin VB.Menu memAudContenedores 
            Caption         =   "Auditar Contenedores"
            HelpContextID   =   2
         End
         Begin VB.Menu memImprimir1raVez 
            Caption         =   "Imprimir Contenedor 1ra Vez"
            HelpContextID   =   2
         End
         Begin VB.Menu memVerContenedores 
            Caption         =   "Ver Contenedores"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menRecepcionMercaderia 
         Caption         =   "Recepcion de Mercaderia"
         HelpContextID   =   1
      End
      Begin VB.Menu menRecepcionIXC 
         Caption         =   "Recepcion de Ingresos X Contabilizar"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu menCatalogosVentas 
         Caption         =   "Catalogos"
         HelpContextID   =   1
         Begin VB.Menu menVerVen 
            Caption         =   "&Vendedores"
            HelpContextID   =   2
         End
         Begin VB.Menu menS7_2 
            Caption         =   "-"
         End
         Begin VB.Menu menDefRangoComi 
            Caption         =   "Definicion de Rangos de Comisiones"
            HelpContextID   =   2
         End
         Begin VB.Menu menTarjetaCredito 
            Caption         =   "Tarjetas de crédito"
            HelpContextID   =   2
         End
         Begin VB.Menu menCupoVendedor 
            Caption         =   "Cupo por Vendedor"
            HelpContextID   =   2
         End
         Begin VB.Menu menCourier 
            Caption         =   "Courier"
            HelpContextID   =   2
         End
         Begin VB.Menu menInsentivos 
            Caption         =   "Incentivos"
         End
      End
      Begin VB.Menu menS7_4 
         Caption         =   "-"
      End
      Begin VB.Menu men_PedBod 
         Caption         =   "&Pedido a Bodega"
         HelpContextID   =   1
      End
      Begin VB.Menu menVerPedBod 
         Caption         =   "&Confirmación de Pedido"
         HelpContextID   =   1
      End
      Begin VB.Menu menVerFacPed 
         Caption         =   "&Despachar Pedidos"
         HelpContextID   =   1
      End
      Begin VB.Menu menNotaEntrega 
         Caption         =   "Nota Entrega"
         HelpContextID   =   1
      End
      Begin VB.Menu menPedidosPendientes 
         Caption         =   "P&edidos Pendientes"
         HelpContextID   =   1
      End
      Begin VB.Menu menFacProVenta 
         Caption         =   "F&acturar Proyecto"
         HelpContextID   =   1
      End
      Begin VB.Menu menManGuiaRemision 
         Caption         =   "&Mantener Guías y Reservas"
         HelpContextID   =   1
      End
      Begin VB.Menu menRegCostoServ 
         Caption         =   "Registro Costos de Servicios"
         HelpContextID   =   1
      End
      Begin VB.Menu menListaEmbarque 
         Caption         =   "Lista de Embarque"
         HelpContextID   =   1
      End
      Begin VB.Menu menRecepcionDePedidos 
         Caption         =   "Recepcion de Pedidos"
         HelpContextID   =   1
      End
      Begin VB.Menu menrManifiestoCarga 
         Caption         =   "Manifiestos de Carga"
         HelpContextID   =   1
      End
      Begin VB.Menu menS7 
         Caption         =   "-"
      End
      Begin VB.Menu menImpresiones 
         Caption         =   "Impresiones"
         HelpContextID   =   1
         Begin VB.Menu menImpCot 
            Caption         =   "&Imprimir Cotización"
            HelpContextID   =   2
         End
         Begin VB.Menu menS7_6 
            Caption         =   "-"
         End
         Begin VB.Menu menReImpFac 
            Caption         =   "&Imprimir Factura"
            HelpContextID   =   2
         End
         Begin VB.Menu menNotRem 
            Caption         =   "Imprimir Notas &Remisión"
            HelpContextID   =   2
         End
         Begin VB.Menu menImpNotaCredito 
            Caption         =   "Imprimir Notas de &Credito"
            HelpContextID   =   2
         End
         Begin VB.Menu menImpIngMer 
            Caption         =   "Imprimir Compra Local/Importación"
            HelpContextID   =   2
         End
         Begin VB.Menu menImprimirEtiquetaDespacho 
            Caption         =   "Imprimir Etiqueta Despacho"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS7_1 
         Caption         =   "-"
      End
      Begin VB.Menu menRanComis 
         Caption         =   "Comisiones y Campo"
         HelpContextID   =   1
         Begin VB.Menu menPordescVenta 
            Caption         =   "Proceso de Comisiones"
            HelpContextID   =   2
         End
         Begin VB.Menu menPorUtiSobreventa 
            Caption         =   "Por utilidad bruta sobre venta"
            HelpContextID   =   2
         End
         Begin VB.Menu menS7_5 
            Caption         =   "-"
         End
         Begin VB.Menu menPorUtiSobreUti 
            Caption         =   "Por utilidad bruta sobre utlidad bruta"
            HelpContextID   =   2
         End
         Begin VB.Menu menActividadCampo 
            Caption         =   "Envio Reporte Campo y Actividad"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menConExistencia 
         Caption         =   "Consulta rapida de Existencias"
         HelpContextID   =   1
      End
      Begin VB.Menu menNotaVenta 
         Caption         =   "Notas de Venta"
         HelpContextID   =   1
      End
      Begin VB.Menu menRimpNV 
         Caption         =   "Reimrimir Nota de Venta"
      End
      Begin VB.Menu menS7_7 
         Caption         =   "-"
      End
      Begin VB.Menu menFacPed 
         Caption         =   "Facturar Pedidos Digitados"
      End
      Begin VB.Menu menVerPedBodFac 
         Caption         =   "Confirmar Facturas Separadas"
      End
   End
   Begin VB.Menu menTesoBanc 
      Caption         =   "&Tesor/Banco"
      Begin VB.Menu menCatalogosTesoBanco 
         Caption         =   "Catalogos"
         HelpContextID   =   1
         Begin VB.Menu menDocCobro 
            Caption         =   "Docu&mentos de Cobro"
            HelpContextID   =   2
         End
         Begin VB.Menu menBanco 
            Caption         =   "&Bancos"
            HelpContextID   =   2
         End
         Begin VB.Menu menCtaBanco 
            Caption         =   "Cta&s. Bancarias"
            HelpContextID   =   2
         End
         Begin VB.Menu menNotaCredito 
            Caption         =   "Notas de C&rédito"
            HelpContextID   =   2
         End
         Begin VB.Menu menNotaDebito 
            Caption         =   "Notas de &Débito"
            HelpContextID   =   2
         End
         Begin VB.Menu menS8 
            Caption         =   "-"
         End
         Begin VB.Menu menEgresoComun 
            Caption         =   "Co&mprobantes de Egreso comunes"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS8_1 
         Caption         =   "-"
      End
      Begin VB.Menu menComprobanteEgreso 
         Caption         =   "&Comprobantes de Egreso"
         HelpContextID   =   1
      End
      Begin VB.Menu menNotCreDeb 
         Caption         =   "&Notas de Crédito y Débito"
         HelpContextID   =   1
      End
      Begin VB.Menu menEstadoCheque 
         Caption         =   "&Estado de Cheques"
         HelpContextID   =   1
      End
      Begin VB.Menu menS8_2 
         Caption         =   "-"
      End
      Begin VB.Menu menCtasCobrar 
         Caption         =   "C&tas. por Cobrar"
         HelpContextID   =   1
      End
      Begin VB.Menu menCobros 
         Caption         =   "Co&bros"
         HelpContextID   =   1
      End
      Begin VB.Menu menEliCobro 
         Caption         =   "Eliminar Cobros"
         HelpContextID   =   1
      End
      Begin VB.Menu menEfecDoc 
         Caption         =   "E&fectivización de Cobros"
         HelpContextID   =   1
      End
      Begin VB.Menu menConfirDocCobrado 
         Caption         =   "Anular Cobros"
         HelpContextID   =   1
      End
      Begin VB.Menu menAplicarCobros 
         Caption         =   "Aplicar Cobros"
         HelpContextID   =   1
      End
      Begin VB.Menu menImpCobros 
         Caption         =   "Imprimir Cobro"
         HelpContextID   =   1
      End
      Begin VB.Menu menCtaCobro 
         Caption         =   "Cta por Cobrar Vendedor"
         HelpContextID   =   1
      End
      Begin VB.Menu menCobrosEfectivizados 
         Caption         =   "Cobros a Efectivizar"
         HelpContextID   =   1
      End
      Begin VB.Menu menModFechaCheque 
         Caption         =   "Modificar Fecha Cheques"
         HelpContextID   =   1
      End
      Begin VB.Menu menModCobros 
         Caption         =   "Modificar Fechas de cobro"
         HelpContextID   =   1
      End
      Begin VB.Menu menArchivoCarteraBanco 
         Caption         =   "Cobranza en Bancos"
         HelpContextID   =   1
      End
      Begin VB.Menu menAplicarNC 
         Caption         =   "Aplicar Notas de Crédito"
         HelpContextID   =   1
      End
      Begin VB.Menu menCobrosAutomaticos 
         Caption         =   "Cobros Automaticos"
         HelpContextID   =   1
      End
      Begin VB.Menu menCargaCobros 
         Caption         =   "Aplicar Cobros Automaticos"
         HelpContextID   =   1
      End
      Begin VB.Menu menAnticipoPedidos 
         Caption         =   "Aplicar Anticipos Pedidos"
         HelpContextID   =   1
      End
      Begin VB.Menu menGenerarArchivoCredito 
         Caption         =   "Archivo de Datos Crediticios"
         HelpContextID   =   1
      End
      Begin VB.Menu menBloqueoDesbloqueo 
         Caption         =   "Bloqueo y Desbloqueo"
         HelpContextID   =   1
      End
      Begin VB.Menu menPermisoDespacho 
         Caption         =   "Autorizar Despachos"
         HelpContextID   =   1
      End
      Begin VB.Menu menS8_3 
         Caption         =   "-"
      End
      Begin VB.Menu menCtasPagar 
         Caption         =   "Cuentas por &Pagar"
         HelpContextID   =   1
      End
      Begin VB.Menu menPagos 
         Caption         =   "Pa&gos"
         HelpContextID   =   1
      End
      Begin VB.Menu menAplicarPagod 
         Caption         =   "Aplicar Pagos"
         HelpContextID   =   1
      End
      Begin VB.Menu menGenerarArchivoPagos 
         Caption         =   "Generar Archivo de Pagos"
         HelpContextID   =   1
      End
      Begin VB.Menu menValidarArchivoPagos 
         Caption         =   "Validar Archivo Pagos"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menImportacion 
      Caption         =   "I&mportación"
      Begin VB.Menu menCatalogosImp 
         Caption         =   "Catalogos"
         HelpContextID   =   1
         Begin VB.Menu menEmbarcador 
            Caption         =   "&Embarcadores"
            HelpContextID   =   2
         End
         Begin VB.Menu menAgente_Afianzado 
            Caption         =   "&Agentes Afianzados"
            HelpContextID   =   2
         End
         Begin VB.Menu menVerificadora 
            Caption         =   "&Verificadoras"
            HelpContextID   =   2
         End
         Begin VB.Menu menS9 
            Caption         =   "-"
         End
         Begin VB.Menu menProceso_Imp 
            Caption         =   "&P&rocesos de Importación"
            HelpContextID   =   2
         End
         Begin VB.Menu menGasto_Importacion 
            Caption         =   "&Gastos de Importación"
            HelpContextID   =   2
         End
      End
      Begin VB.Menu menS9_1 
         Caption         =   "-"
      End
      Begin VB.Menu menIngreso_Imp 
         Caption         =   "&Ingresos de Importación"
         HelpContextID   =   1
      End
      Begin VB.Menu menMod_Imp 
         Caption         =   "Modificar Ingreso Importación"
         HelpContextID   =   1
      End
      Begin VB.Menu menCostoGastos 
         Caption         =   "&Costeo de porductos de Importaciones"
         HelpContextID   =   1
      End
      Begin VB.Menu menPrecioProdImp 
         Caption         =   "Precio de Productos"
         HelpContextID   =   1
      End
      Begin VB.Menu menPedImp 
         Caption         =   "Pedido Importacion"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menCompras 
      Caption         =   "&Compras"
      Begin VB.Menu menComprasCatalogos 
         Caption         =   "Catalogos"
         HelpContextID   =   1
         Begin VB.Menu menPreProducto 
            Caption         =   "Pre Productos"
            HelpContextID   =   2
         End
         Begin VB.Menu menS10_0 
            Caption         =   "-"
         End
      End
      Begin VB.Menu menS10_1 
         Caption         =   "-"
      End
      Begin VB.Menu menVerOrdenCompra 
         Caption         =   "Orden de Compra Inv"
         HelpContextID   =   1
      End
      Begin VB.Menu menVerOrdenCompraSum 
         Caption         =   "Orden de Compra Sum"
         HelpContextID   =   1
      End
   End
   Begin VB.Menu menAT 
      Caption         =   "Ane&xos Transaccionales"
      Begin VB.Menu menAnexos 
         Caption         =   "&Anexos"
         HelpContextID   =   1
      End
      Begin VB.Menu menS12_0 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "mdiPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub memAdminContenedores_Click()
    frmVerContenedorMercaderia.Show
End Sub

Private Sub memAudContenedores_Click()
    With frmVerContenedorMercaderia
        .Show
        .cmdCrearContenedor.Visible = False
        .cmdAnularContenedor.Visible = False
        .cmdReubicarContenedor.Visible = False
        .cmdContenedorVacios.Visible = False
        .cmdTransferirPrendas.Visible = False
        .cmdCrearContenedorDescontinuados.Visible = False
        .cmdImprimirFull.Visible = False
    End With
End Sub

Private Sub memImprimir1raVez_Click()
    With frmVerContenedorMercaderia
        .Show
        .cmdCrearContenedor.Visible = False
        .cmdAuditarContenedo.Visible = False
        .cmdAnularContenedor.Visible = False
        .cmdReubicarContenedor.Visible = False
        .cmdContenedorVacios.Visible = False
        .cmdTransferirPrendas.Visible = False
        .cmdCrearContenedorDescontinuados.Visible = False
        .chkFiltroEstado.Value = 1
        .chkFiltroEstado.Enabled = False
        .cmbEstado.BoundText = "2"
    End With
End Sub

Private Sub memIngContenedores_Click()
    With frmVerContenedorMercaderia
        .Show
        .cmdAnularContenedor.Visible = False
        .cmdReubicarContenedor.Visible = False
        .cmdTransferirPrendas.Visible = False
        .cmdAuditarContenedo.Visible = False
        .cmdContenedorVacios.Visible = False
        .cmdCrearContenedorDescontinuados.Visible = False
        .cmdImprimirFull.Visible = False
    End With
End Sub

Private Sub memIngContenedoresDesc_Click()
    With frmVerContenedorMercaderia
        .Show
        .cmdAnularContenedor.Visible = False
        .cmdReubicarContenedor.Visible = False
        .cmdTransferirPrendas.Visible = False
        .cmdAuditarContenedo.Visible = False
        .cmdContenedorVacios.Visible = False
        .cmdCrearContenedor.Visible = False
        .cmdImprimirFull.Visible = False
    End With

End Sub

Private Sub memVerContenedores_Click()
    With frmVerContenedorMercaderia
        .Show
        .cmdCrearContenedor.Visible = False
        .cmdAnularContenedor.Visible = False
        .cmdReubicarContenedor.Visible = True
        .cmdTransferirPrendas.Visible = True
        .cmdContenedorVacios.Visible = True
        .cmdCrearContenedorDescontinuados.Visible = False
        .cmdImprimirFull.Visible = False
        .cmdAuditarContenedo.Visible = False
    End With
End Sub

Private Sub menActividadCampo_Click()
    frmActividadCampo.Show
End Sub

Private Sub menActualizacionClientes_Click()
    frmClienteModACT.TIPOUsu = "SA"
    frmClienteModACT.Show
End Sub

Private Sub menAdmMovimiento_Click()
    frmVerMovimiento.cmdAnular.Visible = True
    frmVerMovimiento.cmdReasignarContenedores.Visible = True
    frmVerMovimiento.Tag = True
    frmVerMovimiento.Show
End Sub

Private Sub menAplicarNC_Click()
    frmAplicarNC.Show
End Sub

Private Sub menBloqueoDesbloqueo_Click()
    frmBloqueoDesbloqueo.Show
End Sub

Private Sub menCambioProductoPVPUIO_Click()
    frmCambioProducto.Neg = "PDV"
    frmCambioProducto.Show
End Sub

Private Sub menCambioProductoTotal_Click()
    frmCambioProductoTotal.Neg = "%"
    frmCambioProductoTotal.Show
End Sub

Private Sub menCanton_Click()
    frmCanton.Show
End Sub

Private Sub menCargaClientesFidelizacion_Click()
    frmCargaClientesFidelizacion.Show
End Sub

Private Sub menCargaCobros_Click()
    frmCargaCobros.Show
End Sub

Private Sub MDIForm_Load()
    ' Al leer la forma se eliminará las opciones que no son obligatorias
    Limpiar_Menu
    Dim Pasa As Boolean
    Pasa = False
    While Pasa = False
        Pasa = ComprobarFecha
    Wend
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ' Al cerrar la forma se desconectará de la base de datos
    Dim clsRSComp As New clsRegSet
    DesConectar
    clsRSComp.SetRegionalSetting strForNumDec, strForNumMil, strForMonDec, strForMonMil, strForFecha
End Sub


Private Function ComprobarFecha() As Boolean
    Dim clsSql As clsConsulta
    Dim strSql As String
    Dim FechaActual As Date
    Dim FechaClave As Integer
    Dim dias As Integer
    
    FechaActual = Format(Date, "yyyy-mm-dd")

    Set clsSql = New clsConsulta
    clsSql.Inicializar AdoConn, AdoConnMaster

    strSql = " SELECT COALESCE(par_numero,0) as valor " & _
             " FROM parametro " & _
             " WHERE par_codigo='TCL' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
        dias = FormatoD0(clsSql.adorec_Def("valor"))
    Else
        dias = 0
    End If
    
    
     strSql = " SELECT iif(usu_ultimamod is null,0,1) as fecha " & _
             " FROM usuario " & _
             " WHERE usu_codigo='" & strUsuario & "' "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        FechaClave = clsSql.adorec_Def("fecha")
    Else
        FechaClave = 1
    End If
    
    'Compara si debe cambiar
'    If FechaClave = 1 Then
'        If FormatoD0(Format(FechaActual, "dd")) Mod dias = 0 Then
'            ComprobarFecha = False
'            MsgBox "Su contraseña ha caducado. Debe cambiarla este momento!", vbInformation + vbOKOnly, "Cambiar Clave"
'            frmCambioClave.INICIO = True
'            frmCambioClave.Show 1
'        Else
'            ComprobarFecha = True
'        End If
'    Else
        ComprobarFecha = True
'    End If
    Set clsSql = Nothing
End Function


Private Sub men_Click()
Dim xxx As New clsConsulta
    xxx.Inicializar AdoConn, AdoConnMaster
    Dim strSql As String
    Dim ElMenu As String
    Dim CoDiGos() As String
    Dim CodigoMenu As String
    Dim i As Integer
    Dim MaxNiveles As Integer
    
    
    'Borrar el contenido de OP_MENU
    strSql = "DELETE FROM op_menu WHERE op_men_codigo NOT LIKE 'B%'"
    xxx.Ejecutar strSql, "M"
    
    MaxNiveles = 10
    
    ReDim CoDiGos(MaxNiveles)
    CoDiGos(0) = "00"
    Dim LaString As String
    For Each Menu In mdiPrincipal
        'Busca entre los menús cuyo nombre en VB no empiece con una L (Todas las líneas).
        If Menu.Name <> "StatusBar" And InStr(1, Menu.Name, "menS", vbTextCompare) <> 1 And Menu.Visible = True Then
            'Generar código del menú en función del HelpContextID donde manualmente se
            'puso la profundidad del menú
            CoDiGos(Menu.HelpContextID) = Aumentar(CoDiGos(Menu.HelpContextID))
            For i = Menu.HelpContextID + 1 To MaxNiveles
                CoDiGos(i) = "00"
            Next i
            CodigoMenu = ""
            For i = 0 To Menu.HelpContextID
                CodigoMenu = CodigoMenu & CoDiGos(i)
            Next i
            ElMenu = Replace(Menu.Caption, "&", "")
            
            'LaString = LaString & "|" & CodigoMenu & " - " & ElMenu
            strSql = "INSERT INTO op_menu VALUES ('" & CodigoMenu & "','" & ElMenu & "','" & Menu.Name & "','W',0,CURRENT_TIMESTAMP,'" & strUsuario & "') "
            xxx.Ejecutar strSql, "M"
        End If
    Next
    strSql = "DELETE FROM men_permiso "
    xxx.Ejecutar strSql, "M"
    strSql = "INSERT INTO men_permiso SELECT 1,op_men_codigo,CURRENT_TIMESTAMP,'" & strUsuario & "' FROM op_menu "
    xxx.Ejecutar strSql, "M"
    'MsgBox LaString
    MsgBox "Se generó la tabla op_menu", vbInformation, "Información"
End Sub

Private Function Aumentar(Numero As String) As String
    Dim i As Integer
    Dim s As String
    
    i = CInt(Numero)
    i = i + 1
    s = CStr(i)
    While Len(s) < 2
        s = "0" & s
    Wend
    Aumentar = s
End Function


Private Sub men_PedBod_Click()
    frmV_PedBod.Show
End Sub


Private Sub menAcercade_Click()
    'frmBR.webBrowser.Navigate "http://" & strServidorWeb & "/Contabilidad/VerAsiento.php", 2 + 4 + 8, "trabajo"
    'frmVerRecepcionMercaderia.Show
    'frmProductoComisionCampania.Show
    frmEnvioCartera.Show
    'frmImpresionDirecta.Show
End Sub

Private Sub menAdminAF_Click()
    'frmSelActivoFijo.Show
    frmSelAF.Show
End Sub

Private Sub menAdmProducto_Click()
    frmProducto.Show
End Sub


Private Sub menAgente_Afianzado_Click()
    frmSelAgenteAfianzado.Show
End Sub

Private Sub menAjusteInventario_Click()
    frmAjusteInventario.Show
End Sub

Private Sub menAltasAuditoria_Click()
    frmNuevoIngreso.Show
    frmNuevoIngreso.cmbTDoc.BoundText = "AAU"
End Sub

Private Sub menAnexos_Click()
    frmATED.Show
End Sub

Private Sub menAnularCom_Click()
   frmAnularCompra.Show
End Sub

Private Sub menAnularFac_Click()
    frmAnularFactura.Show
End Sub

Private Sub menAplicarCobros_Click()
    frmAplicarCobros.Show
End Sub

Private Sub menAplicarPagod_Click()
    frmAplicarPagos.Show
End Sub

Private Sub menArchivoCarteraBanco_Click()
    frmArchivoCarteraBanco.Show
End Sub


Private Sub menAreasAF_Click()
    frmArea.Show
End Sub


Private Sub menBajasAuditoria_Click()
    frmNuevoEgreso.Show
    frmNuevoEgreso.cmbTDoc.BoundText = "BAU"
End Sub

Private Sub menBanco_Click()
    frmBanco.Show
End Sub

Private Sub menCambioClave_Click()
    frmCambioClave.INICIO = False
    frmCambioClave.Show 1
End Sub

Private Sub menCambioFac_Click()
    frmCambioFac.Show
End Sub

Private Sub menCambioProducto_Click()
    frmCambioProducto.Neg = "%"
    frmCambioProducto.Show
End Sub

Private Sub mencanal_Click()
    frmCanal.Show
End Sub

Private Sub mencascada_Click()
    Me.Arrange 0
End Sub
Private Sub menCatCli_Click()
    frmCatCliente.Show
End Sub

Private Sub menCatPro_Click()
    frmCatProveedor.Show
End Sub

Private Sub menCentroCosto_Click()
    frmCentroCosto.Show
End Sub

Private Sub menCierreMes_Click()
    frmCierreMes.Show
End Sub

Private Sub menCiudad_Click()
    frmCiudad.Show
End Sub

Private Sub menCliente_Click()
    frmClienteMod.TIPOUsu = "AC"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClienteMod_Click()
    frmClienteMod.TIPOUsu = "SU"
    frmClienteMod.cmdReferido.Visible = True
    frmClienteMod.cmdVerReferido.Visible = True
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClienteModCartera_Click()
    frmClienteMod.TIPOUsu = "CA"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = True
    frmClienteMod.Show
End Sub

Private Sub menClienteModGerencia_Click()
    frmClienteMod.TIPOUsu = "GE"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClienteModJefeSAC_Click()
    frmClienteMod.TIPOUsu = "JSA"
    frmClienteMod.cmdReferido.Visible = True
    frmClienteMod.cmdVerReferido.Visible = True
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClienteModSAC_Click()
    frmClienteMod.TIPOUsu = "SAC"
    frmClienteMod.cmdReferido.Visible = True
    frmClienteMod.cmdVerReferido.Visible = True
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClientesPromo_Click()
    frmClienteMod.TIPOUsu = "PR"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClientesPromoGYE_Click()
    frmClienteMod.TIPOUsu = "PRG"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show

End Sub

Private Sub menClientesPVP_Click()
    frmClienteMod.TIPOUsu = "PV"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClientesPVPGYE_Click()
    frmClienteMod.TIPOUsu = "PVG"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menClientesPVPREC_Click()
    frmClienteMod.TIPOUsu = "PVR"
    frmClienteMod.cmdReferido.Visible = False
    frmClienteMod.cmdVerReferido.Visible = False
    frmClienteMod.chkBloqueoDesBloqueo.Visible = False
    frmClienteMod.Show
End Sub

Private Sub menCobrador_Click()
    frmCobrador.Show
End Sub

Private Sub menCobros_Click()
    frmCobros.Show
End Sub

Private Sub menCobrosAutomaticos_Click()
    frmCobrosAutomaticos.Show
End Sub

Private Sub menCobrosEfectivizados_Click()
    frmCobrosEfectivizados.Show
End Sub

Private Sub menColeccion_Click()
    frmColeccion.Show
End Sub

Private Sub menColor_Click()
    frmColor.Show
End Sub

Private Sub menComprasLocales_Click()
    frmNuevoIngreso.Show
    frmNuevoIngreso.cmbTDoc.BoundText = "COM"
End Sub

Private Sub menComprobanteEgreso_Click()
    frmVerComprobanteEgresoComun.Show
End Sub


Private Sub menConExistencia_Click()
    frmConExiPrd.Show
End Sub

Private Sub menConfirDocCobrado_Click()
    frmConfirDocCobrado.Show
End Sub

Private Sub menContabilizarMovMer_Click()
    frmContabilizarMovMer.Show
End Sub

Private Sub menCostoGastos_Click()
    frmGastosCostoImp.Show
End Sub

Private Sub menCourier_Click()
    frmCourier.Show
End Sub

Private Sub menCtaBanco_Click()
    frmCtaBanco.Show
End Sub

Private Sub menCtaCobro_Click()
    frmCtaCobrar.Show
End Sub

Private Sub menCtaMovEgr_Click()
    frmSelMovEgreso.Show
End Sub

Private Sub menCtaMovIng_Click()
    frmSelMovIngreso.Show
End Sub

Private Sub menCtasCobrar_Click()
    frmVerCtaxc_p.Tag = "C"
    frmVerCtaxc_p.Show
End Sub

Private Sub menCtasPagar_Click()
    frmVerCtaxc_p.Tag = "P"
    frmVerCtaxc_p.Show
End Sub

Private Sub menCuadreAsientos_Click()
    frmCuadreAsientos.Show
End Sub

Private Sub menCupoVendedor_Click()
    frmCupoVendedor.Show
End Sub

Private Sub menDefCXC_Click()
    frmDefCXC.Show
End Sub

Private Sub menDefCXP_Click()
    frmDefCXP.Show
End Sub

Private Sub menDefFormaPago_Click()
    frmDefFormaPago.Show
End Sub

Private Sub menDefinicionFlujos_Click()
    frmDefinicionFlujos.Show
End Sub

Private Sub menDefIVAC_Click()
    frmDefIVAC.Show
End Sub

Private Sub menDefIVAV_Click()
    frmDefIVAV.Show
End Sub


Private Sub menDefRangoComi_Click()
    frmV_VerRangoComi.Show
End Sub

Private Sub menDepartamentosAF_Click()
    frmDepartamentos.Show
End Sub

Private Sub menDeposito_Click()
    frmDeposito.Show
End Sub

Private Sub menDepreciacionAF_Click()
    frmVerDepreciacionAF.Show
End Sub

Private Sub menDesmantelar_Click()
    frmDesmantelar.Show
End Sub

Private Sub menDocCobro_Click()
    frmDocPago.Show
End Sub

Private Sub menEfecDoc_Click()
    frmEfecCobro.Show
End Sub

Private Sub menEgresoComun_Click()
    frmSelEgresoComun.Show
End Sub

Private Sub menEliCobro_Click()
    frmEliCobros.Show
End Sub

Private Sub menEmbarcador_Click()
     frmSelEmbarcador.Show
'     frmSelAgenteAfianzado.Show
'     frmSelVerificadora.Show
End Sub

Private Sub menempresa_Click()
    frmSelEmpresa.Show
End Sub

Private Sub menEnvioCorreo_Click()
    frmEnvioCorreo.Show
End Sub

Private Sub menEstadoCheque_Click()
    frmEstadoCheque.Show
End Sub

Private Sub menFacPed_Click()
    frmV_FacPed.Show
End Sub

Private Sub menFacProVenta_Click()
    frmV_FacProVenta.Show
End Sub

Private Sub menFormaEntrega_Click()
    frmFormaEntrega.Show
End Sub

Private Sub menGasto_Importacion_Click()
    frmSelGastoImportacion.Show
End Sub

Private Sub menGenerarArchivoCredito_Click()
    frmGenerarArchivoCredito.Show
End Sub

Private Sub menGenerarArchivoPagos_Click()
    frmGenerarArchivoPagos.Show
End Sub

Private Sub menImpAsiento_Click()
    frmVerAsiento.TipoVisualizacion = False
    frmVerAsiento.Show
End Sub

Private Sub menImpCobros_Click()
    frmV_ReImpCobro.Show
End Sub

Private Sub menImpIngMer_Click()
    frmV_ReImpIngresoMer.Show
End Sub
Private Sub menImprimirEtiquetaDespacho_Click()
    frmImprimirEtiquetaDespacho.Show
End Sub

Private Sub menInsentivos_Click()
    frmIncentivos.Show
End Sub

Private Sub menListaEmbarque_Click()
    frmVerListaEmbarque.Show
    frmVerListaEmbarque.cmdRecibirLista.Visible = False
    frmVerListaEmbarque.cmdNuevo.Visible = True
    frmVerListaEmbarque.cmdCambiarOperador.Visible = True
    frmVerListaEmbarque.cmdImprimirListado.Visible = True
    frmVerListaEmbarque.cmdImprimirEtiqueta.Visible = True
    frmVerListaEmbarque.cmdEnviarCorreo.Visible = True
End Sub

Private Sub menMantGuiaProveedor_Click()
    frmGuiaProveedor.Show
End Sub

Private Sub menImpCompConta_Click()
    frmVerAsiento.TipoVisualizacion = True
    frmVerAsiento.Show
End Sub

Private Sub menImpCot_Click()
    frmV_ImpCotizacion.Show
End Sub

Private Sub menImpNotaCredito_Click()
    frmV_ReImpNotaCredito.Show
End Sub

Private Sub menIngreso_Imp_Click()
    frmVerIngImp.Show
End Sub

Private Sub menLineaN_Click()
    frmLinea.Show
End Sub

Private Sub menLineaP_Click()
    frmLinea.Show
End Sub

Private Sub menListaPrecio_Click()
    frmSelListaPrecio.Show
End Sub

Private Sub menListaProducto_Click()
    frmListaProducto.Show
End Sub

Private Sub menManGuiaRemision_Click()
    frmV_GuiaRemision.Show
End Sub

Private Sub menMarca_Click()
    frmMarca.Show
End Sub

Private Sub menMod_Imp_Click()
    frmModIngImp.Show
End Sub

Private Sub menModCobros_Click()
    frmModCobros.Show
End Sub

Private Sub menModEmpresa_Click()
    frmEmpresa.Show
    frmEmpresa.txtCodigo.Text = strEmpresa
    frmEmpresa.Tag = "M"
End Sub

Private Sub menModFechaCheque_Click()
    frmModCobros.Show
    frmModCobros.dtpFecha.Enabled = False
    frmModCobros.dcmbDocumento.BoundText = "CHP"
    frmModCobros.dcmbDocumento.Locked = True
End Sub

Private Sub menmoshor_Click()
    Me.Arrange 1
End Sub

Private Sub menmosver_Click()
    Me.Arrange 2
End Sub

Private Sub menMotivoAjuste_Click()
    frmMotivoAjuste.Show
End Sub

Private Sub menNotaCredito_Click()
    frmNotaCredito.Show
End Sub

Private Sub menNotaDebito_Click()
    frmNotaDebito.Show
End Sub

Private Sub menNotaEntrega_Click()
    frmV_VerPedConfirm.Show
    frmV_VerPedConfirm.cmdNotaEntrega.Visible = True
    frmV_VerPedConfirm.CmdConfirmar.Visible = False
    frmV_VerPedConfirm.cmdFacturaGuia.Visible = False
    frmV_VerPedConfirm.CmdGuiaRemi.Visible = False
    frmV_VerPedConfirm.cmdPreFactura.Visible = False
End Sub

Private Sub menNotasCreditoCliente_Click()
    frmNuevoIngreso.Show
    frmNuevoIngreso.cmbTDoc.BoundText = "DCL"
End Sub

Private Sub menNotasCreditoProveedor_Click()
    frmNuevoEgreso.Show
    frmNuevoEgreso.cmbTDoc.BoundText = "DPV"
End Sub

Private Sub menNotaVenta_Click()
    frmNuevoEgresoNV.Show
    frmNuevoEgresoNV.cmbTDoc.BoundText = "NOT"
    frmNuevoEgresoNV.dcmbCodP.BoundText = "C00000"
    frmNuevoEgresoNV.CmbFpago.BoundText = "EFE"
    frmNuevoEgresoNV.cmbVendedor.BoundText = "RMC"
End Sub

Private Sub menNotCreDeb_Click()
    frmVerNotasCreditoDebito.Show
End Sub

Private Sub menNotRem_Click()
    frmV_ImpGuiaRemi.Show
End Sub

Private Sub menNuevaEmpresa_Click()
    frmEmpresa.Tag = "N"
    frmEmpresa.Show
End Sub

Public Sub menNuevoBrowser_Click()
    Dim frmBrow As New frmBR
    frmBrow.Show
End Sub

Private Sub menorgico_Click()
    Me.Arrange 3
End Sub

Private Sub menOtroCargo_Click()
    frmOtroCargo.Show
End Sub

Private Sub menPagos_Click()
'    frmCobrosPagos.Tag = "P"
    frmPagos.Show
End Sub

Private Sub menPais_Click()
    frmPais.Show
End Sub

Private Sub menPedidosPendientes_Click()
    frmVerPedPendiente.Show
End Sub

Private Sub menPedImp_Click()
    frmPedImp.Show
End Sub

Private Sub menPermisoDespacho_Click()
    frmPermisoDespacho.Show
End Sub

Private Sub menPlanCuenta_Click()
    frmPlanCuenta.Show
End Sub


Private Sub menPordescVenta_Click()
frmV_VerRangoComi.Show
End Sub

Private Sub menPorUtiSobreUti_Click()
frmV_VerRangoComi3.Show
End Sub

Private Sub menPorUtiSobreventa_Click()
frmV_VerRangoComi2.Show
End Sub

Private Sub menPrecioProdImp_Click()
    frmPrecioProdImp.Show
End Sub

Private Sub menPreProducto_Click()
    frmPreProducto.Show
End Sub

Private Sub menProceso_Imp_Click()
    frmProcesoImportacion.Show
End Sub

Private Sub menProductoClc_Click()
    frmProductoClc.Show
End Sub

Private Sub menProveedor_Click()
    frmProveedor.Show
End Sub

Private Sub menQuitarNotaCredito_Click()
    frmDesaplicarNota.Show
End Sub

Private Sub menRealizaCambioProducto_Click()
    frmRealizaCambioProducto.Neg = "%"
    frmRealizaCambioProducto.Show
End Sub

Private Sub menRealizaCambioProductoPVPUIO_Click()
    frmRealizaCambioProducto.Neg = "PDV"
    frmRealizaCambioProducto.Show
End Sub

Private Sub menRealizaDevolucionProducto_Click()
    frmRealizaDevolucionProducto.Neg = "%"
    frmRealizaDevolucionProducto.Show
End Sub

Private Sub menRealizaDevolucionProductoPVPUIO_Click()
    frmRealizaDevolucionProducto.Neg = "PDV"
    frmRealizaDevolucionProducto.Show
End Sub

Private Sub menRecepcionDePedidos_Click()
    frmVerListaEmbarque.Show
    frmVerListaEmbarque.cmdRecibirLista.Visible = True
    frmVerListaEmbarque.cmdNuevo.Visible = False
    frmVerListaEmbarque.cmdCambiarOperador.Visible = False
    frmVerListaEmbarque.cmdImprimirListado.Visible = True
    frmVerListaEmbarque.cmdImprimirEtiqueta.Visible = False
    frmVerListaEmbarque.cmdEnviarCorreo.Visible = False
End Sub

Private Sub menRecepcionIXC_Click()
    frmRecepcionIXC.Show
End Sub

Private Sub menRecepcionMercaderia_Click()
    frmVerRecepcionMercaderia.Show
End Sub

Private Sub menRegCostoServ_Click()
    frmRegCostosServ.Show
End Sub

Private Sub menRegion_Click()
    frmRegion.Show
End Sub

Private Sub menReImpFac_Click()
    frmV_ReImpFactura.Show
End Sub

Private Sub menReprocesoRet_Click()
    frmReprocesoRet.Show
End Sub

Private Sub menResponsableFlujo_Click()
    frmResponsableFlujo.Show
End Sub

Private Sub menRetenciones_Click()
    frmRetencion.Show
End Sub

Private Sub menRimpNV_Click()
    frmV_ReImpNV.Show
End Sub

Private Sub menrManifiestoCarga_Click()
    frmVerManifiestoCarga.Show
End Sub

Private Sub menSac_Click()
    frmSac.Show
End Sub

Private Sub mensalir_Click()
    Unload Me
End Sub

Private Sub Limpiar_Menu()
    Dim clsCon_Def As clsConsulta
    Dim strSql As String
    Dim strMenu As String
    On Error Resume Next
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' Consulta para conocer los menús que no son obligatorios
        strSql = " SELECT op_men_menu " & _
                 " FROM op_menu " & _
                 " WHERE op_men_obligatorio=0 AND " & _
                       " op_men_plataforma='W' " & _
                 " ORDER BY op_men_codigo "
        clsCon_Def.Ejecutar (strSql)
        clsCon_Def.adorec_Def.MoveFirst
    ' Se esconde todos los menos no obligatorios
        While Not clsCon_Def.adorec_Def.EOF
            strMenu = clsCon_Def.adorec_Def("op_men_menu")
            Me.Controls(strMenu).Visible = False
            clsCon_Def.adorec_Def.MoveNext
        Wend
    Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Public Sub Crear_Menu()
    Dim clsCon_Def As clsConsulta
    Dim strSql As String
    Dim strMenu As String
    On Error Resume Next
        Set clsCon_Def = New clsConsulta
        clsCon_Def.Inicializar AdoConn, AdoConnMaster
    ' Consulta para conocer los menús a los que tiene permiso el usuario de acuerdo al grupo al que pertenece
        strSql = " SELECT distinct op_menu.op_men_menu " & _
                 " FROM ((op_menu INNER JOIN men_permiso ON op_menu.op_men_codigo=men_permiso.op_men_codigo) " & _
                 " INNER JOIN grupo_usuario ON grupo_usuario.gru_u_codigo=men_permiso.gru_u_codigo) " & _
                 " WHERE op_men_obligatorio=0 AND " & _
                       " op_men_plataforma='W' AND " & _
                       " grupo_usuario.usu_codigo= '" & strUsuario & "' "
        clsCon_Def.Ejecutar (strSql)
        If clsCon_Def.adorec_Def.RecordCount > 0 Then
            clsCon_Def.adorec_Def.MoveFirst
    ' Consulta para hacer visibles las opciones a las que se tiene acceso
            While Not clsCon_Def.adorec_Def.EOF
                strMenu = clsCon_Def.adorec_Def("op_men_menu")
                Me.Controls(strMenu).Visible = True
                clsCon_Def.adorec_Def.MoveNext
            Wend
        End If
    Exit Sub
        
errhandler:
    Select Case Err.Number
        Case 1046
            MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                   " not a mysql_real_connect you have to choose a " & vbCrLf & _
                   " database, so Please Choose a database."
        Case Else
            MsgBox "[" & Err.Number & "] " & Err.Description
    End Select
End Sub

Private Sub menSeguimientoFlujo_Click()
    frmSeguimientoFlujo.Show
End Sub

Private Sub menSucursal_Click()
    frmSucursal.Show
End Sub

Private Sub menTalla_Click()
    frmTalla.Show
End Sub

Private Sub menTarjetaCredito_Click()
    frmTarjetaCredito.Show
End Sub

Private Sub menTipDoc_Click()
    frmCatDocumento.Show
End Sub


Private Sub menTipoNegcio_Click()
    frmTipoNegocio.Show
End Sub

Private Sub menTiposAF_Click()
    'frmSelTipoAF.Show
    frmTiposAF.Show
End Sub

Private Sub menTransfBodega_Click()
    frmTransferenciaBod.Show
End Sub

Private Sub menTransformacion_Click()
    frmTransformacion.Show
End Sub


Private Sub menUbicacionBodega_Click()
    frmUbicacionBodega.Show
End Sub

Private Sub menUnidad_Click()
    frmUnidad.Show
End Sub

Private Sub menValidarArchivoPagos_Click()
    frmValidarArchivoPagos.Show
End Sub

Private Sub menVerFacPed_Click()
    frmV_VerPedConfirm.Show
    frmV_VerPedConfirm.cmdNotaEntrega.Visible = False
End Sub

Private Sub menVerGrupos_Click()
    frmVerGrupos.Show
End Sub

Private Sub menVerificadora_Click()
    frmSelVerificadora.Show
End Sub

Private Sub menVerIngInventario_Click()
    frmVerIngInventario.Show
End Sub

Private Sub menVerMovimiento_Click()
    frmVerMovimiento.cmdAnular.Visible = False
    frmVerMovimiento.cmdReasignarContenedores.Visible = False
    frmVerMovimiento.Tag = False
    frmVerMovimiento.Show
End Sub

Private Sub menVerOrdenCompra_Click()
    frmVerOrdenCompra.strTipoOrdenCompra = "P"
    frmVerOrdenCompra.Show
End Sub

Private Sub menVerOrdenCompraSum_Click()
    frmVerOrdenCompra.strTipoOrdenCompra = "S"
    frmVerOrdenCompra.Show
End Sub

Private Sub menVerPedBod_Click()
    frmV_VerPedBod.Show
End Sub

Private Sub menVerPedBodFac_Click()
    frmV_VerPedBodFac.Show
End Sub

Private Sub menVerProductos_Click()
    frmVerProducto.Show
End Sub

Private Sub menVerVen_Click()
    frmVendedor.Show
End Sub

Private Sub menZona_Click()
    frmZona.Show
End Sub

Private Sub manGuiasProveedor_Click()
    frmNuevoIngreso.Show
    frmNuevoIngreso.cmbTDoc.BoundText = "IGR"
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrePromCot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################################
'#  Esta clase realiza consultas contra la base de datos para generar al final
'#  una tabla temporal con los precios promedio de los productos que intervienen
'#  en una cotización
'#
'#  clsPrePromCot V1.0
'#  Copyright (C) 2002
'#
'#  Procesos internos que maneja:
'#  *   Consulta los datos de una determinada contización, asignada a un cliente
'#      específico de una empresa.
'#  *   Consulta a su vez los productos simples que son parte de los productos
'#      compuestos, para saber en que porcentaje intervienen en el precio total
'#      del producto compuesto cotizado.
'#  *   Luego se calculan las precios los productos antes mensionados para obtener
'#      su precio con relación al porcentaje que ellos intervienen en el producto
'#      compuesto.
'#  *   Con estos datos se crea un tabla temporal que contiene los productos
'#      simples que son parte de los productos compuestos y sus respectivos precios.
'#  *   Luego se consultan los productos simples que intervienen en la cotización
'#      con sus precios, para insertarlos en la tabla temporal antes mensionada.
'#  *   Finalmente se hace un promedio del precio por producto de la tabla
'#      temporal y se crea una nueva tabla temporal con estos datos.
'#
'#  Tablas que maneja:
'#
'#  lista_precio_p:
'#  *   De aquí se consultan los precios actuales de los productos para calcular
'#      su porcentaje de intervención en el producto compuesto.
'#  cotizacion:
'#  *   De aquí se consultan los datos de los productos tanto simples como compuestos
'#      que intervienen en una cotización.
'#  producto:
'#  *   Con esta tabla se cruzan las demás tablas para conocer los productos simples.
'#  producto_compuesto:
'#  *   Con esta tabla se cruzan las demás tablas para conocer los productos compuestos.
'#  PrePromCot:
'#  *   Tabla en donde se almacenan los precios promedios de la cotización
'#
'################################################################################

Private clsSql As New clsConsulta
Private strSQL As String

Public Sub crearTabla(codiCot As Variant, codiPer As String, codiEmp As String)
    If codiCot = "" Or codiPer = "" Or codiEmp = "" Then
            frmV_PedBod.Controlado = False
        Exit Sub
    End If
    'Crea una tabla temporal con el valor real de cada producto compuesto de la cotización seleccionada
    strSQL = " CREATE TABLE #AuxiPreComp " & _
             " SELECT producto_compuesto.prd_com_codigo, Sum(det_prd_com_cantidad*lis_pre_p_precio) AS preComp, lista_precio_p.lis_pre_codigo " & _
             " FROM ((((((cotizacion INNER JOIN det_cotizacion ON (cotizacion.emp_codigo = det_cotizacion.emp_codigo) AND (cotizacion.cot_codigo = det_cotizacion.cot_codigo)) " & _
             " INNER JOIN producto_compuesto ON (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo) AND (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo)) " & _
             " INNER JOIN det_prd_com ON (producto_compuesto.prd_com_codigo = det_prd_com.prd_com_codigo) AND (producto_compuesto.emp_codigo = det_prd_com.emp_codigo)) " & _
             " INNER JOIN producto ON (det_prd_com.emp_codigo = producto.emp_codigo) AND (det_prd_com.prd_codigo = producto.prd_codigo)) " & _
             " INNER JOIN lista_precio_p ON (producto.prd_codigo = lista_precio_p.prd_codigo) AND (producto.emp_codigo = lista_precio_p.emp_codigo)) " & _
             " INNER JOIN categoria_p ON (lista_precio_p.lis_pre_codigo = categoria_p.lis_pre_codigo) AND (lista_precio_p.emp_codigo = categoria_p.emp_codigo)) " & _
             " INNER JOIN persona ON (categoria_p.emp_codigo = persona.emp_codigo) AND (categoria_p.cat_p_codigo = persona.cat_p_codigo) AND (categoria_p.cat_p_tipo = persona.cat_p_tipo) " & _
             " Where cotizacion.cot_codigo='" & codiCot & "' And persona.per_codigo='" & codiPer & "' AND cotizacion.emp_codigo='" & codiEmp & "'" & _
             " GROUP BY producto_compuesto.prd_com_codigo, categoria_p.lis_pre_codigo, cotizacion.emp_codigo " & _
             " ORDER BY producto_compuesto.prd_com_codigo "
    clsSql.Ejecutar strSQL
    'Crea una tabla temporal con el porcentaje que aporta cada producto simple de un producto compuesto con respecto al precio total del producto compuesto
    strSQL = " CREATE TABLE #AuxiPorcComp " & _
             " SELECT producto_compuesto.prd_com_codigo, det_prd_com.prd_codigo, producto.prd_nombre, lis_pre_p_precio/preComp AS Porcent " & _
             " FROM (((#AuxiPreComp INNER JOIN producto_compuesto ON #AuxiPreComp.prd_com_codigo = producto_compuesto.prd_com_codigo) " & _
             " INNER JOIN det_prd_com ON (producto_compuesto.prd_com_codigo = det_prd_com.prd_com_codigo) AND (producto_compuesto.emp_codigo = det_prd_com.emp_codigo)) " & _
             " INNER JOIN producto ON (det_prd_com.prd_codigo = producto.prd_codigo) AND (det_prd_com.emp_codigo = producto.emp_codigo)) " & _
             " INNER JOIN lista_precio_p ON (#AuxiPreComp.lis_pre_codigo = lista_precio_p.lis_pre_codigo) AND (producto.emp_codigo = lista_precio_p.emp_codigo) AND (producto.prd_codigo = lista_precio_p.prd_codigo) " & _
             " ORDER BY producto_compuesto.prd_com_codigo "
    clsSql.Ejecutar strSQL
    'Obtiene los precios reales a partir de la cotización, de los productos simples que son parte de los productos compuestos
    strSQL = " CREATE TABLE #AuxiPrdsFac " & _
             " SELECT det_prd_com.prd_codigo, AuxiPorcComp.prd_nombre, (det_cot_cantidad*det_prd_com_cantidad) as cant, ROUND(Porcent*det_cot_precio,4) AS Precio " & _
             " FROM (((cotizacion INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
             " INNER JOIN producto_compuesto ON (det_cotizacion.prd_codigo = producto_compuesto.prd_com_codigo) AND (det_cotizacion.emp_codigo = producto_compuesto.emp_codigo)) " & _
             " INNER JOIN det_prd_com ON (producto_compuesto.emp_codigo = det_prd_com.emp_codigo) AND (producto_compuesto.prd_com_codigo = det_prd_com.prd_com_codigo)) " & _
             " INNER JOIN #AuxiPorcComp ON (det_prd_com.prd_codigo = #AuxiPorcComp.prd_codigo) AND (#AuxiPorcComp.prd_com_codigo = det_cotizacion.prd_codigo) " & _
             " WHERE cotizacion.emp_codigo='" & codiEmp & "' AND cotizacion.cot_codigo='" & codiCot & "' " & _
             " ORDER BY det_prd_com.prd_codigo "
    clsSql.Ejecutar strSQL
    'Inserta en la tabla temporal AuxiPrdsFac los restantes productos simples cotizados, que no son parte de los productos compuestos
    strSQL = " Insert INTO #AuxiPrdsFac " & _
             " SELECT producto.prd_codigo, producto.prd_nombre, det_cotizacion.det_cot_cantidad, det_cotizacion.det_cot_precio " & _
             " FROM ((cotizacion INNER JOIN det_cotizacion ON (cotizacion.cot_codigo = det_cotizacion.cot_codigo) AND (cotizacion.emp_codigo = det_cotizacion.emp_codigo)) " & _
             " INNER JOIN producto ON (det_cotizacion.prd_codigo = producto.prd_codigo) AND (det_cotizacion.emp_codigo = producto.emp_codigo)) " & _
             " WHERE cotizacion.cot_codigo='" & codiCot & "' AND cotizacion.emp_codigo='" & codiEmp & "'"
    clsSql.Ejecutar strSQL
'******* PRECIOS PROMEDIOS DE PRODUCTOS
    'Elimina la tabla de precios promedios si existe
    elimTabla
    'Hace una agrupación y calcula el promedio de los productos que intervienen en una cotización
    strSQL = " CREATE TABLE #PrePromCot " & _
             " SELECT prd_codigo, Sum(cant) AS cant, ROUND(Sum(Precio*cant)/Sum(cant),4) AS PromPre " & _
             " From #AuxiPrdsFac " & _
             " GROUP BY prd_codigo, prd_nombre " & _
             " ORDER BY prd_codigo "
    clsSql.Ejecutar strSQL
    'Borra las tablas temporales creadas
    strSQL = " EXEC Sp_Drop_Table_if_Exist '#AuxiPorcComp' "
    clsSql.Ejecutar strSQL
    strSQL = " EXEC Sp_Drop_Table_if_Exist '#AuxiPrdsFac' "
    clsSql.Ejecutar strSQL
    strSQL = " EXEC Sp_Drop_Table_if_Exist '#AuxiPreComp' "
    clsSql.Ejecutar strSQL
    frmV_PedBod.Controlado = True
End Sub

Public Sub elimTabla()
    'Elimina la tabla de precios promedios si existe
    strSQL = " EXEC Sp_Drop_Table_if_Exist '#PrePromCot' "
    clsSql.Ejecutar strSQL
End Sub

Public Function nombTabla() As String
    'Devuelve el nombre de la tabla temporal de precios promedios de una cotización
    nombTabla = "#PrePromCot"
End Function

Private Sub Class_Initialize()
    'Inicializa el objeto de consulta a la base de datos
    clsSql.Inicializar AdoConn, AdoConnMaster
End Sub

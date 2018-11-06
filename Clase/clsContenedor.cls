VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContenedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private clsCon_Def As New clsConsulta
Public strNumContenedor As String
Public strFecha As String
Public intEstado As Integer
Public strBodega As String
Public strUbicacion As String
Public strObservacion As String
Private strSql As String
Public adorec_DetalleContenedor As New ADODB.Recordset

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub

Public Sub setVacio()
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strFecha = ""
    strBodega = ""
    strUbicacion = ""
    strObservacion = ""
    Set adorec_DetalleContenedor.DataSource = Nothing
End Sub

Public Function SetContenedor(strContenedor As String) As Boolean
    Dim strSql As String
    Dim clsAux As New clsConsulta
'    AdoConn.Close
'
'        AdoConn.ConnectionTimeout = 120
'        AdoConn.CommandTimeout = 120
'        AdoConn.CursorLocation = adUseClient
'    AdoConn.Open
'    AdoConnMaster.Close
'    AdoConnMaster.Open
    clsAux.Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT con_mer_fecha,dep_codigo,ubi_bod_codigo,est_con_mer_codigo,con_mer_observacion " & _
             " FROM contenedor_mercaderia " & _
             " WHERE emp_codigo = '" & strEmpresa & "' " & _
             " AND con_mer_codigo LIKE  '" & strContenedor & "'"
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        strNumContenedor = strContenedor
        strFecha = clsAux.adorec_Def("con_mer_fecha")
        strBodega = clsAux.adorec_Def("dep_codigo")
        intEstado = clsAux.adorec_Def("est_con_mer_codigo")
        strUbicacion = clsAux.adorec_Def("ubi_bod_codigo")
        strObservacion = clsAux.adorec_Def("con_mer_observacion")
        strSql = " SELECT producto.prd_codigo,prd_nombre,SUM(IIF(con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as tot " & _
                 " FROM det_contenedor_mercaderia INNER JOIN producto " & _
                 " ON det_contenedor_mercaderia.emp_codigo=producto.emp_codigo " & _
                 " AND det_contenedor_mercaderia.prd_codigo=producto.prd_codigo" & _
                 " WHERE det_contenedor_mercaderia.emp_codigo = '" & strEmpresa & "' " & _
                 " AND det_contenedor_mercaderia.con_mer_codigo LIKE  '" & strContenedor & "'" & _
                 " AND det_contenedor_mercaderia.det_con_mer_cantidad!=0 " & _
                 " GROUP BY producto.prd_codigo,prd_nombre " & _
                 " ORDER BY tot DESC,prd_nombre "
        clsAux.Ejecutar strSql
        Set adorec_DetalleContenedor.DataSource = clsAux.adorec_Def.DataSource
        SetContenedor = True
    Else
        SetContenedor = False
    End If
    Set clsAux = Nothing
End Function

Public Sub NuevoContenedor(strSuc1 As String, strPto1 As String, Fecha As String, estado As Integer, bodega As String, ubicacion As String, Observacion As String, Optional TipoColor As Variant = 0, Optional TipoPrenda As Variant = 0, Optional EsDeInventario As Boolean = False, Optional OrdenCompra As String = "")
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strFecha = Fecha
    intEstado = estado
    strBodega = bodega
    strUbicacion = ubicacion
    strSql = "BEGIN TRAN "
    clsAux.Ejecutar strSql, "M"
    strNumContenedor = NuevoNumeroContenedor(EsDeInventario)
    If EsDeInventario = True Then
        Observacion = UCase(Observacion) & vbNewLine & "CONTENEDOR DE INVENTARIO"
    End If
    strSql = " INSERT INTO contenedor_mercaderia (emp_codigo, con_mer_codigo,tip_con_codigo,tip_mer_con_codigo, con_mer_fecha,est_con_mer_codigo, " & _
             " dep_codigo, ubi_bod_codigo, con_mer_observacion, con_mer_fechamod, con_mer_usumod,ord_com_codigo)" & _
             " VALUES('" & strEmpresa & "','" & strNumContenedor & "','" & TipoColor & "','" & TipoPrenda & "','" & strFecha & "','" & intEstado & "'," & _
             " '" & strBodega & "','" & strUbicacion & "','" & UCase(Observacion) & "',CURRENT_TIMESTAMP,'" & strUsuario & "','" & FormatoD0(OrdenCompra) & "')"
    clsAux.Ejecutar strSql, "M"
    strSql = "COMMIT TRAN "
    clsAux.Ejecutar strSql, "M"
    strSql = " INSERT INTO contenedor_mercaderia_historia (emp_codigo,con_mer_codigo,con_mer_his_fecha, " & _
             " dep_codigo, ubi_bod_codigo, con_mer_his_observacion," & _
             " con_mer_his_fechamod,con_mer_his_usumod)" & _
             " VALUES('" & strEmpresa & "','" & strNumContenedor & "','" & strFecha & "'," & _
             " '" & strBodega & "','" & strUbicacion & "','" & UCase(Observacion) & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Sub

Public Sub AgregarDetalle(producto As String, ContenedorOrigen As String, ContenedorDestino As String, Fecha As String, Cantidad As Double, Optional tipoMov As String = "", Optional numeroMov As String = "0")

    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    If ContenedorOrigen = "" Then ContenedorOrigen = 0
    If ContenedorDestino = "" Then ContenedorDestino = 0
    strSql = " SELECT COUNT(*) as n " & _
             " FROM det_contenedor_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_mer_codigo='" & strNumContenedor & "'" & _
             " AND prd_codigo='" & producto & "'" & _
             " AND con_mer_codigo_origen='" & ContenedorOrigen & "'" & _
             " AND con_mer_codigo_destino='" & ContenedorDestino & "'" & _
             " AND tip_mov_codigo='" & tipoMov & "'" & _
             " AND mov_codigo='" & FormatoD0(numeroMov) & "'" & _
             " AND LEFT(det_con_mer_fecha,10)='" & Left(Fecha, 10) & "'"
    clsAux.Ejecutar strSql, "M"
    If FormatoD0(clsAux.adorec_Def("n")) > 0 Then
        ModificaDetContenedor producto, ContenedorOrigen, ContenedorDestino, Fecha, Cantidad, tipoMov, FormatoD0(numeroMov)
    Else
        NuevoDetContenedor producto, ContenedorOrigen, ContenedorDestino, Fecha, Cantidad, tipoMov, FormatoD0(numeroMov)
    End If
    
    Set clsAux = Nothing
End Sub

Public Sub AgregarDetalleCaracteristica(Tipo As String, Numero As Double, Observacion As String, Imagen As String)
    Dim clsAux As New clsConsulta
    Dim rs As New ADODB.Recordset
    Dim mystream As ADODB.Stream
    
    clsAux.Inicializar AdoConn, AdoConnMaster
    If Tipo = "" Or (Numero = 0 And Observacion = "") Then Exit Sub
    
    strSql = " INSERT INTO contenedor_mercaderia_caracteristica " & _
             " (emp_codigo, con_mer_codigo, con_mer_car_fecha, con_mer_car_tipo, con_mer_car_numero, con_mer_car_observacion, con_mer_car_fechamod, con_mer_car_usumod)" & _
             " VALUES('" & strEmpresa & "', '" & strNumContenedor & "',CURRENT_TIMESTAMP,'" & Tipo & "','" & Numero & "','" & Observacion & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar strSql, "M"
    
    If Imagen <> "" Then
        Set mystream = New ADODB.Stream
        mystream.Type = adTypeBinary
        rs.Open " SELECT con_mer_car_foto " & _
                " FROM contenedor_mercaderia_caracteristica " & _
                " WHERE emp_codigo='" & strEmpresa & "'" & _
                " AND con_mer_codigo='" & strNumContenedor & "' " & _
                " AND con_mer_car_tipo='" & Tipo & "' " & _
                " AND con_mer_car_numero='" & FormatoD4(Numero) & "' " & _
                " AND con_mer_car_observacion='" & Observacion & "'", AdoConn, adOpenKeyset, adLockOptimistic
        mystream.Open
        mystream.LoadFromFile Imagen
        rs!con_mer_car_foto = mystream.Read
        rs.Update
        mystream.Close
        rs.Close
    End If
    
    Set clsAux = Nothing
End Sub

Private Sub ModificaDetContenedor(producto As String, ContenedorOrigen As String, ContenedorDestino As String, Fecha As String, Cantidad As Double, tipoMov As String, numeroMov As String)
    strSql = " UPDATE det_contenedor_mercaderia " & _
             " SET det_con_mer_cantidad=det_con_mer_cantidad+'" & Cantidad & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_mer_codigo='" & strNumContenedor & "'" & _
             " AND prd_codigo='" & producto & "'" & _
             " AND con_mer_codigo_origen='" & ContenedorOrigen & "'" & _
             " AND con_mer_codigo_destino='" & ContenedorDestino & "'" & _
             " AND tip_mov_codigo='" & tipoMov & "'" & _
             " AND mov_codigo='" & numeroMov & "'" & _
             " AND LEFT(det_con_mer_fecha,10)='" & Left(Fecha, 10) & "'"
    clsCon_Def.Ejecutar strSql, "M"
End Sub

Public Sub AgregaObservacion(Observacion As String)
    strSql = " UPDATE contenedor_mercaderia " & _
             " SET con_mer_observacion=CONCAT('" & UCase(Observacion) & vbNewLine & "',con_mer_observacion)" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_mer_codigo='" & strNumContenedor & "'"
    clsCon_Def.Ejecutar strSql, "M"
End Sub

Public Sub CambiaEstado(estado As Integer)
    strSql = " UPDATE contenedor_mercaderia " & _
             " SET est_con_mer_codigo='" & estado & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_mer_codigo='" & strNumContenedor & "'"
    clsCon_Def.Ejecutar strSql, "M"
End Sub

Private Sub NuevoDetContenedor(producto As String, ContenedorOrigen As String, ContenedorDestino As String, Fecha As String, Cantidad As Double, tipoMov As String, numeroMov As String)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " INSERT INTO det_contenedor_mercaderia (emp_codigo, con_mer_codigo, prd_codigo,con_mer_codigo_origen, con_mer_codigo_destino," & _
             " det_con_mer_fecha,tip_mov_codigo,mov_codigo,det_con_mer_cantidad, det_con_mer_fechamod, det_con_mer_usumod)" & _
             " VALUES('" & strEmpresa & "','" & strNumContenedor & "','" & producto & "','" & ContenedorOrigen & "','" & ContenedorDestino & "', " & _
             " '" & Fecha & "','" & tipoMov & "','" & numeroMov & "','" & Cantidad & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Sub

Private Function NuevoNumeroContenedor(Optional EsDeInventario As Boolean = False) As String
        Dim clsAux As New clsConsulta
    Dim ano As String
    Dim Mes As String
    If EsDeInventario = False Then
        ano = Format(Right(Format(strFecha, "YYYY"), 2), "00")
        Mes = Format(Format(strFecha, "mm"), "00")
    Else
        ano = "18"
        Mes = "00"
    End If
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT COALESCE(MAX(RIGHT(con_mer_codigo,7)),0)+1 as n " & _
             " FROM contenedor_mercaderia WITH (TABLOCKX)" & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND con_mer_codigo LIKE '" & ano & Mes & "%'"
    clsAux.Ejecutar strSql, "M"
    If clsAux.adorec_Def.RecordCount > 0 Then
        NuevoNumeroContenedor = ano & Mes & Format(clsAux.adorec_Def("n"), "0000000")
    Else
        NuevoNumeroContenedor = ano & Mes & Format(1, "0000000")
    End If
    Set clsAux = Nothing
End Function

Public Sub ImprimirSTK(Optional ImprimirDirecto As Boolean = True)
    Dim Stk As New frmReporte
    
    If ImpresoraEtiqueta = "" Then
        Stk.VSPrint.PrintDialog pdPrint
        ImpresoraEtiqueta = Stk.VSPrint.Device
        GuardarImpresoras
    End If
    Stk.VSPrint.Device = ImpresoraEtiqueta
    Stk.VSPrint.PaperWidth = 7669.292
    Stk.VSPrint.PaperHeight = 3885.039
    
    Stk.strNumero = strNumContenedor
    Stk.strReporte = "rptSTKContenedorMercaderia"
    Stk.Show
    Stk.Form_Activate
    If ImprimirDirecto = True Then
        Stk.VSPrint.Copies = 1
        Stk.VSPrint.PrintDoc
        Unload Stk
    End If
    
End Sub

Public Sub ImprimirLista(Optional Copias As Integer = 1, Optional ImprimirDirecto As Boolean = True)
    Dim Lista As New frmReporte
    Lista.strNumero = strNumContenedor
    Lista.strReporte = "rptContenedorMercaderia"
    Lista.Show
    Lista.Form_Activate
    If ImprimirDirecto = True Then
        Lista.VSPrint.Copies = Copias
        Lista.VSPrint.PrintDoc
        Unload Lista
    End If
End Sub
Public Sub AnularContenedor(Optional Motivo As String = "")
    Dim anula As Boolean
    strSql = " SELECT count(*) as n " & _
             " FROM det_ingreso_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND con_mer_codigo LIKE '" & strNumContenedor & "'"
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def("n") > 0 Then
        MsgBox "Contenedor asignado a un ingreso de mercaderia no puede ser anulado", vbInformation, "Inventario"
    Else
        anula = True
        While Motivo = ""
            Motivo = InputBox("Motivo de Anulacion", "Inventario")
            If Motivo = "" Then
                If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el contenedor?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                    anula = False
                    Motivo = "NO ANULAR"
                End If
            Else
                anula = True
            End If
        Wend
        Motivo = Motivo & vbNewLine & strUsuario & vbNewLine & HoyDia & " " & Format(Ahora, "HH:MM:SS")
        If anula = True Then
            strSql = " UPDATE contenedor_mercaderia  " & _
                     " SET est_con_mer_codigo=-1, " & _
                     " con_mer_observacion=CONCAT('" & UCase(Motivo) & "',' - ',con_mer_observacion)," & _
                     " con_mer_fechamod=CURRENT_TIMESTAMP," & _
                     " con_mer_usumod='" & strUsuario & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND con_mer_codigo='" & strNumContenedor & "'"
            clsCon_Def.Ejecutar strSql, "M"
        End If
    End If
End Sub

Public Sub AnularMovimiento(strTipo As String, strNumero As String)
    Dim anula As Boolean
    anula = True
    If strTipo = "COM" Or strTipo = "IIM" Or strTipo = "ITN" Then
        If MsgBox("Desea liberar la recepcion para volver a ingresar?", vbQuestion + vbYesNo, "Anular de Contenedor") = vbYes Then
            anula = False
        End If
    End If
    If anula = True Then
        strSql = " UPDATE det_contenedor_mercaderia " & _
                 " SET det_con_mer_cantidad=0," & _
                 " det_con_mer_fechamod=CURRENT_TIMESTAMP, " & _
                 " det_con_mer_usumod='" & strUsuario & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND tip_mov_codigo LIKE '" & strTipo & "'" & _
                 " AND mov_codigo LIKE '" & strNumero & "'"
    Else
        strSql = " UPDATE det_contenedor_mercaderia " & _
                 " SET mov_codigo=0," & _
                 " tip_mov_codigo='' " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND tip_mov_codigo LIKE '" & strTipo & "'" & _
                 " AND mov_codigo LIKE '" & strNumero & "'"
    End If
    clsCon_Def.Ejecutar strSql, "M"
    
End Sub

Public Sub EgresarPrendas(producto As String, Cantidad As Double, bodega As String, tipoMov As String, numeroMov As String)
    Dim clsAux As New clsConsulta
    Dim saldo As Double
    Dim cantEgre As Double
    Dim strContenedorDef As String
    Dim Contenedor As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    If strContenedorRecurrente = "" Then
        strContenedorDef = ""
    Else
        
        While Contenedor = "" Or strContenedorRecurrente = "111"
            If Contenedor = "" And strContenedorRecurrente = "111" Then
                Contenedor = InputBox("Ingrese el numero de contenedor en el que quiere sacar la referencia", "Contenedor", "")
            Else
                Contenedor = strContenedorRecurrente
            End If
            strSql = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as exi" & _
                     " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                     " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                     " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                     " INNER JOIN ubicacion_bodega ON contenedor_mercaderia.emp_codigo=ubicacion_bodega.emp_codigo " & _
                     " AND contenedor_mercaderia.ubi_bod_codigo=ubicacion_bodega.ubi_bod_codigo AND contenedor_mercaderia.dep_codigo=ubicacion_bodega.dep_codigo " & _
                     " WHERE contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                     " AND contenedor_mercaderia.est_con_mer_codigo>=0 " & _
                     " AND contenedor_mercaderia.dep_codigo in ('" & bodega & "') " & _
                     " AND contenedor_mercaderia.con_mer_codigo='" & Contenedor & "' " & _
                     " AND det_contenedor_mercaderia.prd_codigo='" & producto & "' " & _
                     " GROUP BY contenedor_mercaderia.con_mer_codigo,ubi_bod_facturable " & _
                     " HAVING SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)>0 " & _
                     " ORDER BY ubi_bod_facturable ASC, exi ASC"
            clsAux.Ejecutar strSql
            If clsAux.adorec_Def.RecordCount = 0 Then
                Contenedor = ""
            Else
                If clsAux.adorec_Def("exi") < Cantidad Then
                    Contenedor = ""
                Else
                    strContenedorRecurrente = Contenedor
                End If
            End If
            If Contenedor = "" Then MsgBox "Ingrese un contenedor de la bodega " & bodega & " y que tenga el producto " & producto, vbCritical, "Contenedor"
        Wend
        strContenedorDef = " AND contenedor_mercaderia.con_mer_codigo='" & Contenedor & "' "
        
    End If
    
    strSql = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as exi" & _
             " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
             " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
             " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
             " INNER JOIN ubicacion_bodega ON contenedor_mercaderia.emp_codigo=ubicacion_bodega.emp_codigo " & _
             " AND contenedor_mercaderia.ubi_bod_codigo=ubicacion_bodega.ubi_bod_codigo AND contenedor_mercaderia.dep_codigo=ubicacion_bodega.dep_codigo " & _
             " WHERE contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
             " AND contenedor_mercaderia.est_con_mer_codigo>=0 " & strContenedorDef & _
             " AND contenedor_mercaderia.dep_codigo IN('" & bodega & "') " & _
             " AND det_contenedor_mercaderia.prd_codigo='" & producto & "' " & _
             " GROUP BY contenedor_mercaderia.con_mer_codigo,ubi_bod_facturable " & _
             " HAVING SUM(IIF(det_contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)>0 " & _
             " ORDER BY ubi_bod_facturable ASC, exi ASC"
    clsAux.Ejecutar strSql
    saldo = Cantidad
    HoyDiaHora = Ahora
    If clsAux.adorec_Def.RecordCount <= 0 Then
        RegistraError "clsContenedor", "Funcion EgresarPrendas(" & producto & "," & Cantidad & "," & bodega & "," & tipoMov & "," & numeroMov & ") clsAux.adorec_Def.RecordCount <= 0", strSql
    End If
    While (Not clsAux.adorec_Def.EOF) And saldo > 0
        If clsAux.adorec_Def("exi") >= saldo Then
            cantEgre = saldo
            saldo = 0
        Else
            cantEgre = clsAux.adorec_Def("exi")
            saldo = saldo - clsAux.adorec_Def("exi")
        End If
        SetContenedor clsAux.adorec_Def("con_mer_codigo")
        AgregarDetalle producto, clsAux.adorec_Def("con_mer_codigo"), "", HoyDiaHora, FormatoD2(cantEgre), tipoMov, numeroMov
        clsAux.adorec_Def.MoveNext
        
    Wend
    'controlarMov tipoMov, numeroMov
    Set clsAux = Nothing
End Sub

Private Sub controlarMov(tipoMov As String, numeroMov As String)
    Dim clsAux As clsConsulta
    Dim cant As Double
    Dim SQL1 As String
    Dim SQL2 As String
    Set clsAux = New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    SQL1 = "select sum(det_egr_cantidad) as aa " & _
           " FROM det_egreso where tip_egr_codigo='" & tipoMov & "' and egr_codigo='" & numeroMov & "'"
    clsAux.Ejecutar SQL1, "M"
    cant = clsAux.adorec_Def("aa")
    SQL2 = "select sum(det_con_mer_cantidad) as aa " & _
           " from det_contenedor_mercaderia where tip_mov_codigo='" & tipoMov & "' and mov_codigo='" & numeroMov & "'"
    clsAux.Ejecutar SQL2, "M"
    If cant <> clsAux.adorec_Def("aa") Then
        RegistraError "clsContenedor", "Funcion controlarMov(" & tipoMov & "," & numeroMov & ") cant <> clsAux.adorec_Def(aa)", SQL1 & "; " & SQL2
    End If
End Sub
Public Sub CambiarDocumentoPrendas(producto As String, tipoMov As String, numeroMov As String, tipoMovNuevo As String, numeroMovNuevo As String)
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " UPDATE det_contenedor_mercaderia " & _
             " SET tip_mov_codigo='" & tipoMovNuevo & "', " & _
             " mov_codigo='" & numeroMovNuevo & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_mov_codigo='" & tipoMov & "' " & _
             " AND mov_codigo='" & numeroMov & "' " & _
             " AND prd_codigo='" & producto & "'"
    clsAux.Ejecutar strSql
End Sub


Public Sub IngresarPrendas(producto As String, Cantidad As Double, bodega As String, tipoMov As String, numeroMov As String)
    Dim clsAux As New clsConsulta
    Dim clsConte As New clsContenedor
    Dim Contenedor As String
    Dim ubicacion As String
    Dim strPregunta As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    HoyDiaHora = Ahora
        strSql = " SELECT contenedor_mercaderia.con_mer_codigo,SUM(IIF(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad) as exi" & _
                 " FROM contenedor_mercaderia INNER JOIN det_contenedor_mercaderia " & _
                 " ON contenedor_mercaderia.emp_codigo=det_contenedor_mercaderia.emp_codigo " & _
                 " AND contenedor_mercaderia.con_mer_codigo=det_contenedor_mercaderia.con_mer_codigo " & _
                 " INNER JOIN ubicacion_bodega ON contenedor_mercaderia.emp_codigo=ubicacion_bodega.emp_codigo " & _
                 " AND contenedor_mercaderia.ubi_bod_codigo=ubicacion_bodega.ubi_bod_codigo AND contenedor_mercaderia.dep_codigo=ubicacion_bodega.dep_codigo " & _
                 " WHERE contenedor_mercaderia.emp_codigo='" & strEmpresa & "' " & _
                 " AND contenedor_mercaderia.est_con_mer_codigo>=0 " & _
                 " AND contenedor_mercaderia.dep_codigo='" & bodega & "' " & _
                 " AND det_contenedor_mercaderia.prd_codigo='" & producto & "' " & _
                 " GROUP BY contenedor_mercaderia.con_mer_codigo,ubi_bod_facturable " & _
                 " HAVING SUM(IIF(contenedor_mercaderia.con_mer_codigo=con_mer_codigo_origen,-1,1)*det_con_mer_cantidad)>0" & _
                 " ORDER BY ubi_bod_facturable ASC, exi DESC"
        clsAux.Ejecutar strSql

    If strContenedorRecurrente = "" Then
        strPregunta = "No encontro ningún contenedor con esta referencia en la bodega seleccionada." & vbNewLine & _
                      "Desea ingresar el contenedor en el que se guardará?" & vbNewLine & _
                      "Si responde que NO se creará un contenedor nuevo"
    Else
        strPregunta = "Desea ingresar el contenedor en el que se guardará?" & vbNewLine & _
                      "Si responde que NO se creará un contenedor nuevo"
    End If
    If clsAux.adorec_Def.RecordCount > 0 And strContenedorRecurrente = "" Then
        Contenedor = clsAux.adorec_Def("con_mer_codigo")
        SetContenedor Contenedor
    ElseIf strContenedorRecurrente <> "" And strContenedorRecurrente <> "111" Then
        Contenedor = strContenedorRecurrente
        SetContenedor Contenedor
    Else
        If booUnContenedor = False Then
            If MsgBox(strPregunta, vbQuestion + vbYesNo, "Contenedor") = vbYes Then
                While Contenedor = ""
                    Contenedor = InputBox("Ingrese el numero de contenedor en el que quiere guardar la referencia", "Contenedor", "")
                    strSql = " SELECT count(*) as n " & _
                             " FROM contenedor_mercaderia " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND dep_codigo='" & bodega & "' " & _
                             " AND con_mer_codigo='" & Contenedor & "' "
                    clsAux.Ejecutar strSql
                    If clsAux.adorec_Def.RecordCount = 0 Then
                        Contenedor = ""
                    Else
                        If clsAux.adorec_Def("n") = 0 Then
                            Contenedor = ""
                        End If
                    End If
                    If Contenedor = "" Then MsgBox "Ingrese un contenedor de la bodega " & bodega, vbCritical, "Contenedor"
                Wend
                SetContenedor Contenedor
                strContenedorRecurrente = Contenedor
                booUnContenedor = True
            Else
                clsConte.Inicializar AdoConn, AdoConnMaster
                
                strSql = " SELECT ubi_bod_codigo " & _
                         " FROM ubicacion_bodega " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND dep_codigo='" & bodega & "' "
                clsAux.Ejecutar strSql
                If clsAux.adorec_Def.RecordCount = 1 Then
                    ubicacion = clsAux.adorec_Def("ubi_bod_codigo")
                Else
                    ubicacion = ""
                End If
                While ubicacion = ""
                    ubicacion = InputBox("Ingrese la ubicacion del contenedor en el que quiere guardar la referencia", "Contenedor", "")
                    strSql = " SELECT count(*) as n " & _
                             " FROM ubicacion_bodega " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND dep_codigo='" & bodega & "' " & _
                             " AND ubi_bod_codigo='" & ubicacion & "' "
                    clsAux.Ejecutar strSql
                    If clsAux.adorec_Def.RecordCount = 0 Then
                        ubicacion = ""
                    Else
                        If clsAux.adorec_Def("n") = 0 Then
                            ubicacion = ""
                        End If
                    End If
                    If ubicacion = "" Then MsgBox "Ingrese una ubicacion valida para la bodega " & bodega, vbCritical, "Contenedor"
                Wend
                
                clsConte.NuevoContenedor strSucursal, strPtoFactura, Ahora, 1, bodega, ubicacion, "Creado por que no hay contenedores con esta referencia"
                Contenedor = clsConte.strNumContenedor
                MsgBox "Contenedor creado " & Contenedor
                If strContenedorRecurrente <> "" Then
                    booUnContenedor = True
                    strContenedorRecurrente = Contenedor
                ElseIf MsgBox("Desea que todas las referencias de la transaccion se vayan al mismo contenedor?", vbYesNo + vbQuestion, "Contenedores") = vbYes Then
                    booUnContenedor = True
                    strContenedorRecurrente = Contenedor
                Else
                    booUnContenedor = False
                End If
                SetContenedor Contenedor
            End If
        Else
            Contenedor = strContenedorRecurrente
            SetContenedor Contenedor
        End If
    End If
    AgregarDetalle producto, "", Contenedor, HoyDiaHora, FormatoD2(Cantidad), tipoMov, numeroMov
End Sub

Public Sub TransferirPrendasA(Productos() As String, clsContenedorDestino As clsContenedor)
    Dim i As Long
    Dim ProdCant() As String
    Dim Observacion As String
    Dim booContinuar As Boolean
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim frmEntrega As New frmReporte
    
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    booContinuar = False
    If clsContenedorDestino.strBodega = strBodega Then
        booContinuar = True
    Else
        If MsgBox("¿Va hacer una transferencia del mercaderia entre bodegas?", vbYesNo, "Trasnferencia") = vbYes Then
            'trasferencia entre bodegas
            Observacion = "TRASNFERENCIA CONTENEDOR " & strNumContenedor & vbNewLine & _
                          "de la bodega: " & strBodega & " - a la bodega: " & clsContenedorDestino.strBodega & vbNewLine & _
                          vbNewLine & Observacion
            clsEgreso.NuevoEgr False, "ETR", False, strSucursal, strPtoFactura, , , , HoyDia, , , Observacion
            clsIngreso.NuevoIng False, "ITR", False, strSucursal, strPtoFactura, , , , HoyDia, , , Observacion

            booContinuar = True
        End If
    End If
    HoyDiaHora = Ahora
    If booContinuar = True Then
    
        For i = 1 To UBound(Productos)
            ProdCant = Split(Productos(i), vbTab)
            If clsContenedorDestino.strBodega <> strBodega Then
                clsEgreso.NuevoDetEgr ProdCant(0), strBodega, FormatoD2(ProdCant(1))
                clsIngreso.NuevoDetIng ProdCant(0), clsContenedorDestino.strBodega, FormatoD2(ProdCant(1))
            End If

            AgregarDetalle ProdCant(0), strNumContenedor, clsContenedorDestino.strNumContenedor, HoyDiaHora, FormatoD2(ProdCant(1)), clsEgreso.strTipo, clsEgreso.strDoc
            clsContenedorDestino.AgregarDetalle ProdCant(0), strNumContenedor, clsContenedorDestino.strNumContenedor, HoyDiaHora, FormatoD2(ProdCant(1)), clsIngreso.strTipo, clsIngreso.strDoc
        Next i
        InicializarContenedorRecurrente
    End If
    
    If clsContenedorDestino.strBodega <> strBodega Then

        frmEntrega.strNumero = clsEgreso.strDoc
        frmEntrega.strTipo = clsEgreso.strTipo
        frmEntrega.strReporte = "rptTransferencia"
        frmEntrega.Show
        
    End If
    
    Set clsEgreso = Nothing
    Set clsIngreso = Nothing
    
End Sub

Public Sub CambiarUbicacionContenedor(NuevaBodega As String, NuevaUbicacion As String, Observacion As String, Optional SinTransferencia As Boolean = False)
    Dim booContinuar As Boolean
    booContinuar = False
    If SinTransferencia = False Then
        If NuevaBodega = strBodega Then
            booContinuar = True
        Else
            If MsgBox("¿Va hacer una transferencia del mercaderia entre bodegas?", vbYesNo, "Trasnferencia") = vbYes Then
                'trasferencia entre bodegas
                strContenedorRecurrente = strNumContenedor
                TrasnferirContenedor strBodega, NuevaBodega, Observacion
                InicializarContenedorRecurrente
                booContinuar = True
            End If
        End If
    Else
        booContinuar = True
        Observacion = "Sin Transferencia - " & Observacion
    End If
    If booContinuar = True Then
        strSql = " UPDATE contenedor_mercaderia  " & _
                 " SET dep_codigo='" & NuevaBodega & "', " & _
                 " ubi_bod_codigo='" & NuevaUbicacion & "', " & _
                 " con_mer_observacion=CONCAT('" & IIf(Len(Trim(Observacion)) > 0, UCase(Trim(Observacion)) & vbNewLine, "") & "',con_mer_observacion)," & _
                 " con_mer_fechamod=CURRENT_TIMESTAMP," & _
                 " con_mer_usumod='" & strUsuario & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND con_mer_codigo='" & strNumContenedor & "'"
        clsCon_Def.Ejecutar strSql, "M"
        strSql = " INSERT INTO contenedor_mercaderia_historia (emp_codigo,con_mer_codigo,con_mer_his_fecha, " & _
                 " dep_codigo, ubi_bod_codigo, con_mer_his_observacion," & _
                 " con_mer_his_fechamod,con_mer_his_usumod)" & _
                 " VALUES('" & strEmpresa & "','" & strNumContenedor & "',CURRENT_TIMESTAMP," & _
                 " '" & NuevaBodega & "','" & NuevaUbicacion & "','" & IIf(Len(Trim(Observacion)) > 0, UCase(Trim(Observacion)) & vbNewLine, "") & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
        strBodega = NuevaBodega
        strUbicacion = NuevaUbicacion
        strObservacion = IIf(Len(Trim(Observacion)) > 0, UCase(Trim(Observacion)) & vbNewLine, "") & strObservacion
    End If
End Sub


Private Sub TrasnferirContenedor(BodegaOrigen As String, BodegaDestino As String, Observacion As String)
    Dim clsIngreso As New clsInventario
    Dim clsEgreso As New clsInventario
    Dim frmEntrega As New frmReporte
    clsIngreso.Inicializar AdoConn, AdoConnMaster
    clsEgreso.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT prd_codigo,det_con_mer_cantidad" & _
             " FROM det_contenedor_mercaderia " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND con_mer_codigo='" & strNumContenedor & "'"
    'clsCon_Def.Ejecutar strSql
    adorec_DetalleContenedor.MoveFirst
    If adorec_DetalleContenedor.RecordCount > 0 Then
        Observacion = "TRASNFERENCIA CONTENEDOR " & strNumContenedor & vbNewLine & _
                      "de la bodega: " & BodegaOrigen & " - a la bodega: " & BodegaDestino & vbNewLine & _
                      vbNewLine & Observacion
        clsEgreso.NuevoEgr True, "ETR", False, strSucursal, strPtoFactura, , , , HoyDia, , , Observacion
        clsIngreso.NuevoIng True, "ITR", False, strSucursal, strPtoFactura, , , , HoyDia, , , Observacion
        
        While Not adorec_DetalleContenedor.EOF
            clsEgreso.NuevoDetEgr adorec_DetalleContenedor("prd_codigo"), BodegaOrigen, adorec_DetalleContenedor("tot")
            clsIngreso.NuevoDetIng adorec_DetalleContenedor("prd_codigo"), BodegaDestino, adorec_DetalleContenedor("tot")
            adorec_DetalleContenedor.MoveNext
        Wend
        InicializarContenedorRecurrente
    End If
    
    frmEntrega.strNumero = clsEgreso.strDoc
    frmEntrega.strTipo = clsEgreso.strTipo
    frmEntrega.strReporte = "rptTransferencia"
    frmEntrega.Show
    
    Set clsEgreso = Nothing
    Set clsIngreso = Nothing
    
End Sub

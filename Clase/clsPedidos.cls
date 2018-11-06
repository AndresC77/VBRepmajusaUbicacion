VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private clsCon_Def As clsConsulta
Public strPedido As String
Public strTipoPedido As String
Public dblSubtotalMasIva As Double
Private strSql As String


Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
    
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub

Public Sub RecalculoTotal(Optional strPed As String = "")
    If strPed <> "" Then
        strPedido = strPed
    End If
    
    strSql = " SELECT pedido.ped_codigo, " & _
             " ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-ped_dctoadicional/100.00),2)" & _
             " + ROUND((SUM((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - SUM(IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(1-ped_dctoadicional/100.00) * (par_numero)/100.00,2) as tot" & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
             " AND pedido.per_codigo=persona.per_codigo " & _
             " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo AND det_ped_incentivo=0  " & _
             " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo" & _
             " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
             " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo " & _
             " AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
             " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
             " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
             " AND pedido.ped_codigo=producto_promo2.ped_codigo" & _
             " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
             " AND pedido.ped_codigo='" & strPedido & "'" & _
             " GROUP BY pedido.ped_codigo,pedido.ped_dctoadicional,parametro.par_numero"
    clsCon_Def.Ejecutar strSql, "M"
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        dblSubtotalMasIva = FormatoD2(clsCon_Def.adorec_Def("tot"))
    Else
        dblSubtotalMasIva = 0
    End If
    strSql = " UPDATE pedido " & _
             " SET ped_subtotal='" & FormatoD2(dblSubtotalMasIva) & "' " & _
             " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
             " AND pedido.ped_codigo='" & strPedido & "'"
    clsCon_Def.Ejecutar strSql, "M"
End Sub

Public Sub Eliminar(Optional strPed As String = "")
    If strPed <> "" Then
        strPedido = strPed
    End If
    strSql = " DELETE FROM det_pedido " & _
             " WHERE det_pedido.emp_codigo = '" & strEmpresa & "' " & _
             " AND det_pedido.ped_codigo='" & strPedido & "'"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " DELETE FROM pedido " & _
             " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
             " AND pedido.ped_codigo='" & strPedido & "'"
    clsCon_Def.Ejecutar strSql, "M"
End Sub

Public Function GenerarReprogramacion(Optional strPed As String = "") As String
    Dim num As String
    Dim PedRetorno As String
    Dim clsAuxPed As New clsConsulta
    clsAuxPed.Inicializar AdoConn, AdoConnMaster
    If strPed <> "" Then
        strPedido = strPed
    End If
    
    PedRetorno = ""
    
    strSql = " SELECT exi_fut_fecha,sum(det_ped_cant_programada) as cant " & _
             " FROM pedido INNER JOIN det_pedido " & _
             " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
             " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
             " INNER JOIN existencia_futura ON existencia_futura.emp_codigo=det_pedido.emp_codigo" & _
             " AND existencia_futura.prd_codigo=det_pedido.prd_codigo" & _
             " AND existencia_futura.exi_fut_fecha>pedido.ped_fecha " & _
             " Where pedido.emp_codigo='" & strEmpresa & "' " & _
             " AND pedido.ped_codigo = '" & strPedido & "'" & _
             " GROUP BY exi_fut_fecha "
    clsAuxPed.Ejecutar strSql
    While Not clsAuxPed.adorec_Def.EOF
        strSql = " BEGIN TRAN "
        clsCon_Def.Ejecutar strSql, "M"
        strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0(Left(strPedido, Len(strPedido) - 7) & "0000001") & "') as num " & _
                 " From pedido WITH (TABLOCKX) " & _
                 " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(Left(strPedido, Len(strPedido) - 7)) & "%'" & _
                 " GROUP BY emp_codigo"
        clsCon_Def.Ejecutar (strSql), "M"
        num = clsCon_Def.adorec_Def("num")
        strSql = " INSERT INTO pedido " & _
                 " SELECT emp_codigo ,'" & num & "', per_codigo, ven_codigo, '', '0', tipo_fac_codigo, " & _
                 " null, tar_cre_codigo, tar_cre_porcentaje, ped_fecha, -4, 0, " & _
                 " CONCAT('Ped.Reprogramado: " & strPedido & " - Para despachar:" & clsAuxPed.adorec_Def("exi_fut_fecha") & " - ',ped_observacion), " & _
                 " ped_egr_bodega, ped_direccion_envio, '', '', '', ped_dctoadicional, " & _
                 " CURRENT_TIMESTAMP, ped_usumod,1 " & _
                 " FROM pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo='" & strPedido & "'"
        clsCon_Def.Ejecutar (strSql), "M"
        strSql = " COMMIT TRAN "
        clsCon_Def.Ejecutar strSql, "M"
        
        strSql = " INSERT INTO det_pedido " & _
                 " SELECT pedido.emp_codigo, '" & num & "', det_pedido.prd_codigo, dep_codigo, det_ped_cant_programada, det_ped_cant_programada," & _
                 " 0, det_ped_precio, det_ped_dcto/(det_ped_cant_entregada+det_ped_cant_programada)*det_ped_cant_programada, det_ped_descripcion, CURRENT_TIMESTAMP," & _
                 " det_ped_usumod, det_ped_incentivo, 0,0" & _
                 " FROM pedido INNER JOIN det_pedido " & _
                 " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN existencia_futura ON existencia_futura.emp_codigo=det_pedido.emp_codigo" & _
                 " AND existencia_futura.prd_codigo=det_pedido.prd_codigo" & _
                 " AND existencia_futura.exi_fut_fecha='" & clsAuxPed.adorec_Def("exi_fut_fecha") & "' " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' AND pedido.ped_codigo='" & strPedido & "' " & _
                 " AND det_ped_cant_programada!=0"
        clsCon_Def.Ejecutar strSql, "M"
        
        strSql = " UPDATE det_pedido " & _
                 " SET det_ped_ped_reprogramado='" & num & "'" & _
                 " FROM det_pedido INNER JOIN existencia_futura ON existencia_futura.emp_codigo=det_pedido.emp_codigo" & _
                 " AND existencia_futura.prd_codigo=det_pedido.prd_codigo" & _
                 " AND existencia_futura.exi_fut_fecha='" & clsAuxPed.adorec_Def("exi_fut_fecha") & "'" & _
                 " WHERE det_pedido.emp_codigo='" & strEmpresa & "' AND det_pedido.ped_codigo='" & strPedido & "' " & _
                 " AND det_ped_cant_programada!=0"
        clsCon_Def.Ejecutar strSql, "M"
        RecalculoTotal num
        PedRetorno = PedRetorno & num & ","
        clsAuxPed.adorec_Def.MoveNext
        If strPed <> "" Then
            strPedido = strPed
        End If
    Wend
    If PedRetorno <> "" Then
        PedRetorno = Left(PedRetorno, Len(PedRetorno) - 1)
    End If
    GenerarReprogramacion = PedRetorno
End Function

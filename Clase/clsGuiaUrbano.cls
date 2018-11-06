VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuiaUrbano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Req As WinHttp.WinHttpRequest
Attribute Req.VB_VarHelpID = -1
Dim strParametros As String

Public Function Generar(NumeroContenedor As String) As String
    Dim strUsuario As String
    Dim strClave As String
    Dim strHttp As String
    Dim Resultado() As String
    Dim ResultadoIndividual() As String
    Dim strMensajeRes As String
    Dim strGuia As String
    Dim strFecha As String
    Dim strSQL As String
    Dim datFecha As Date
    Dim jj As Long
    Dim Detalle As String
    Dim clsAux As New clsConsulta
    
    clsAux.Inicializar AdoConn, AdoConnMaster
    strFecha = Format(datFecha, "dd/MM/yyyy")
    
'    strUsuario = "general-shi2351"
'    strClave = "7c222fb2927d828af22f592134e8932480637c0d"
'    strHttp = "http://200.105.232.250:8181/ws/ue/guia_electronica/"
    strUsuario = "2669-WS"
    strClave = "7c222fb2927d828af22f592134e8932480637c0d"
    strHttp = "https://app.urbano.com.ec/ws/ue/guia_electronica/"
    
    
    strSQL = " SELECT DISTINCT '1' as n,det_contenedor.emp_codigo,det_contenedor.con_codigo,CAST(pedido.ped_codigo as varchar) as obs,ped_direccion_envio,pedido.per_codigo " & _
             " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
             " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
             " AND det_contenedor.tip_egr_codigo=pedido.ped_tip_egr_codigo " & _
             " AND pedido.ped_estado=2 " & _
             " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & NumeroContenedor & "' " & _
             " UNION " & _
             " SELECT '2' as n,det_contenedor_per.emp_codigo,det_contenedor_per.con_codigo,det_contenedor_per.det_con_per_detalle,'','' " & _
             " FROM det_contenedor_per " & _
             " WHERE det_contenedor_per.emp_codigo='" & strEmpresa & "' AND det_contenedor_per.con_codigo='" & NumeroContenedor & "' " & _
             " ORDER BY n,per_codigo"
    clsAux.Ejecutar strSQL
    Detalle = ""
    jj = 0
    While Not clsAux.adorec_Def.EOF
        Detalle = Detalle & clsAux.adorec_Def("obs") & ","
        jj = jj + 1
        clsAux.adorec_Def.MoveNext
    Wend
        
        
    strSQL = " SELECT ncajas.tc,det_con_caj_peso, contenedor.con_codigo, FORMAT(con_fecha,'dd/MM/yyyy') as con_fecha, con_guia,con_peso, " & _
             " '" & jj & "' AS numped, " & _
             " CONCAT(per_apellido,' ',per_nombre) as per,per_ruc,RTRIM(SUBSTRING(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio),CHARINDEX(' - ',IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio))+3,500)) as per_direccion, " & _
             " RTRIM(LEFT(IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio),CHARINDEX(' - ',IIF(d.ped_direccion_envio='' OR d.ped_direccion_envio IS NULL OR LEFT(d.ped_direccion_envio,8)='DIRECTOR',CONCAT(ciu_nombre,'/',can_nombre,'/',pai_nombre,' - ',per_direccion),d.ped_direccion_envio)))) as per_direccion1,per_direccion2,CONCAT(per_telf,'/',per_fax) as per_telf,per_celular,per_email,ciu_nombre,ciu_ubigeos,can_nombre, " & _
             " CAST(REPLACE('" & Detalle & "',',',', ') AS varchar) AS detalle " & _
             " FROM empresa INNER JOIN contenedor ON empresa.emp_codigo=contenedor.emp_codigo " & _
             " INNER JOIN courier ON contenedor.emp_codigo=courier.emp_codigo AND contenedor.cou_codigo=courier.cou_codigo " & _
             " INNER JOIN persona ON empresa.emp_codigo=persona.emp_codigo " & _
             " AND contenedor.per_codigo=persona.per_codigo " & _
             " INNER JOIN ciudad ON persona.ciu_codigo=ciudad.ciu_codigo " & _
             " INNER JOIN canton ON ciudad.can_codigo=canton.can_codigo " & _
             " INNER JOIN pais ON ciudad.pai_codigo=pais.pai_codigo "
    strSQL = strSQL & " INNER JOIN det_contenedor_caja ON contenedor.emp_codigo=det_contenedor_caja.emp_codigo " & _
             " AND contenedor.con_codigo=det_contenedor_caja.con_codigo " & _
             " INNER JOIN paquete_envio ON det_contenedor_caja.emp_codigo=paquete_envio.emp_codigo " & _
             " AND det_contenedor_caja.paq_env_codigo=paquete_envio.paq_env_codigo " & _
             " INNER JOIN (SELECT emp_codigo,con_codigo,count(*) as tc FROM det_contenedor_caja " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND con_codigo='" & NumeroContenedor & "' " & _
             " GROUP BY emp_codigo,con_codigo) ncajas " & _
             " ON contenedor.emp_codigo=ncajas.emp_codigo" & _
             " and contenedor.con_codigo=ncajas.con_codigo"
    strSQL = strSQL & " LEFT JOIN (SELECT TOP 1 pedido.emp_codigo,pedido.per_codigo,pedido.ped_direccion_envio " & _
             " FROM det_contenedor INNER JOIN pedido ON det_contenedor.emp_codigo=pedido.emp_codigo " & _
             " AND det_contenedor.egr_codigo=pedido.ped_egr_codigo " & _
             " AND det_contenedor.tip_egr_codigo=pedido.ped_tip_egr_codigo " & _
             " AND pedido.ped_estado=2 " & _
             " INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo " & _
             " AND pedido.per_codigo=persona.per_codigo " & _
             " AND persona.for_pag_codigo IN ('CONT','EFE')" & _
             " WHERE det_contenedor.emp_codigo='" & strEmpresa & "' AND det_contenedor.con_codigo='" & NumeroContenedor & "' ) d" & _
             " ON persona.emp_codigo=d.emp_codigo " & _
             " AND persona.per_codigo=d.per_codigo "
    strSQL = strSQL & " WHERE empresa.emp_codigo='" & strEmpresa & "' AND contenedor.con_codigo='" & NumeroContenedor & "' "
    clsAux.Ejecutar strSQL
    strFecha = clsAux.adorec_Def("con_fecha")
strParametros = "{" & "'vp_linea':'" & 3 & "'," & _
                "'vp_cod_rastreo':'" & NumeroContenedor & "'," & _
                "'vp_cod_barra':'" & NumeroContenedor & "'," & _
                "'vp_fech_emi_vent':'" & strFecha & "'," & _
                "'vp_nro_o_compra':''," & _
                "'vp_nro_guia_trans':''," & _
                "'vp_nro_factura':''," & _
                "'vp_cod_cliente':'" & clsAux.adorec_Def("per_ruc") & "'," & _
                "'vp_nom_cliente':'" & clsAux.adorec_Def("per") & "','vp_nom_empresa':'" & clsAux.adorec_Def("per") & "'," & _
                "'vp_nro_telf':'" & clsAux.adorec_Def("per_telf") & "','vp_nro_telf_mobil':'" & clsAux.adorec_Def("per_celular") & "', " & _
                "'vp_correo_elec':'" & clsAux.adorec_Def("per_email") & "'," & _
                "'vp_dir_entrega':'" & clsAux.adorec_Def("per_direccion") & "'," & _
                "'vp_nro_via':'0', " & _
                "'vp_nro_int':'','vp_nom_urb':'','vp_ubi_direc':'" & clsAux.adorec_Def("ciu_ubigeos") & "','vp_ref_direc':''," & _
                "'vp_id_direc':'0','vp_fech_pro':'','vp_arco_hor':''," & _
                "'vp_fech_venc':''," & _
                "'vp_nom_autorizado':'','vp_nro_doc_autorizado':''," & _
                "'vp_nom_autorizado_2':'','vp_nro_doc_autorizado_2':''," & _
                "'vp_med_pago':'','vp_descripcion':'','vp_anotacion':'','vp_moneda':'DOL','vp_importe':''," & _
                "'vp_serv_recol':'NO'," & _
                "'vp_cod_cen_act':'','vp_nom_cen_act':'','vp_dir_cen_act':'','vp_ubi_cen_act':''," & _
                "'vp_rec_arco_hor':''," & _
                "'vp_peso_total':'" & clsAux.adorec_Def("con_peso") & "'," & _
                "'vp_pieza_total':'" & clsAux.adorec_Def("tc") & "','vp_urgente':'NO','vp_picking':'NO','vp_mecanizado':'NO','vp_venta_seller':'NO'," & _
                "'vp_sell_codigo':'','vp_sell_nombre':'','vp_sell_direcc':'','vp_sell_ubigeo':''}"

strParametros = Replace(strParametros, "'", """")

Set Req = New WinHttp.WinHttpRequest
    With Req
        .Open "POST", strHttp, False
        .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .setRequestHeader "user", strUsuario
        .setRequestHeader "pass", strClave
        
        .send "json=" & strParametros
        Resultado = Split(Replace(Replace(.responseText, "{", ""), "}", ""), ",")
        ResultadoIndividual = Split(Resultado(1), ":")
        strMensajeRes = Replace(ResultadoIndividual(1), """", "")
        If strMensajeRes = "OK" Then
            ResultadoIndividual = Split(Resultado(2), ":")
            strGuia = Replace(ResultadoIndividual(1), """", "")
            MsgBox "GUIA: " & strGuia
            strSQL = " UPDATE contenedor " & _
                     " SET con_guia='" & strGuia & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND con_codigo='" & NumeroContenedor & "'"
            clsAux.Ejecutar strSQL, "M"
            Generar = strGuia
        Else
            MsgBox "ERROR: " & strMensajeRes
            Generar = "NO GENERA"
        End If
    End With
End Function

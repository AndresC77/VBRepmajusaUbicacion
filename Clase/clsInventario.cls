VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsCon_Def As clsConsulta
Public strTipo As String
Public strDoc As String
Public strPedido As String
Public strFecha As String
Public strIE As String
Public strPersona As String
Public strGuia As String
Public strGuiaCodigo As String
Public strIXC As String
Public dblSubTotal As Double
Public dblSubTotal0 As Double
Public dblDcto As Double
Public dblIVA As Double
Public dblTotal As Double
Public dblRetencion As Double
Public dblTotalProd As Double
Public dblTotalServ As Double
Public dblTotalProdIVA As Double
Public dblTotalServIVA As Double
Public booSecPublico As Boolean
Public booSinIva As Boolean
Public IngresaAutoAContenedor As Boolean
Private clsCostear_Def As clsCostear

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
    
    Set clsCostear_Def = New clsCostear
    clsCostear_Def.Inicializar adocon_ParL, adocon_ParM
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub

Public Sub CrearGuia(strCourier As String, strPlaca As String, strSuc As String, strPto As String)
    Dim strSql As String
    Dim strGuiaCodigo As String
    strSql = " BEGIN TRAN"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " SELECT COALESCE(MAX(egr_gui_codigo),0)+1 AS num " & _
                 " FROM egreso_guia WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql, "M"
    strGuiaCodigo = clsCon_Def.adorec_Def("num")
    strSql = " SELECT COALESCE(MAX(egr_gui_numero),0)+1 AS num " & _
                 " FROM egreso_guia WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_gui_serie ='" & strPto & strSuc & "'"
    clsCon_Def.Ejecutar strSql, "M"
    strGuia = clsCon_Def.adorec_Def("num")
    strSql = " INSERT INTO egreso_guia (emp_codigo,egr_gui_codigo, tip_egr_codigo, egr_codigo, cou_codigo, " & _
             " egr_gui_serie, egr_gui_numero, egr_gui_placa, egr_gui_fechamod, egr_gui_usumod) " & _
             " VALUES('" & strEmpresa & "','" & strGuiaCodigo & "','" & strTipo & "','" & strDoc & "','" & strCourier & "'," & _
             " '" & strPto & strSuc & "','" & strGuia & "','" & strPlaca & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsCon_Def.Ejecutar strSql, "M"
    strSql = " COMMIT TRAN"
    clsCon_Def.Ejecutar strSql, "M"
    DocElectronico "06", strGuiaCodigo
End Sub

Public Function NuevoIng(IngAutoAContenedor As Boolean, Tipo As String, ConfirmaNumero As Boolean, Optional strSuc As String = "", Optional strPto As String = "", Optional strNumero As String = "", Optional ForPag As String = "", Optional Persona As String = "", Optional Fecha As String = "", Optional Doc2 As String = "", Optional Vendedor As String = "", Optional Observ As String = "", Optional Asiento As String = "", Optional strAutor As String = "", Optional strCaduca As String = "00/0000", Optional Subtotal As Double = 0, Optional SubTotal_0 As Double = 0, Optional Dcto As Double = 0, Optional Impto As Double = 0, Optional Total As Double = 0, Optional saldo As Double = 0, Optional SecPublico As Boolean = False, Optional SinIVA As Boolean = False, Optional codIVA As Integer = 0, Optional IXC As String = "") As Boolean
    Dim strSql As String
    Dim Doc As String
    Dim NumIng As Double
    Dim DebeEntrar As Boolean
    strIXC = IXC
    DebeEntrar = False
    booSecPublico = SecPublico
    booSinIva = SinIVA
    Observ = UCase(Observ)
    IngresaAutoAContenedor = IngAutoAContenedor
    strIE = "I"
    strTipo = Tipo
    If Fecha = "" Then
        strFecha = Format(HoyDia, "yyyy-mm-dd")
    Else
        strFecha = Fecha
    End If
    NuevoIng = False
    
    strSql = " BEGIN TRAN "
    clsCon_Def.Ejecutar strSql, "M"
    
    If Tipo = "COM" Then
        If NumeroDeNuevoDocIng(ConfirmaNumero, strSucursal, strPtoFacturaOriginal, Doc) = True Then
            DebeEntrar = True
            NumIng = strSucursal & strPtoFacturaOriginal & Format(Doc, "0000000")
        End If
    Else
        If NumeroDeNuevoDocIng(ConfirmaNumero, strSuc, strPto, Doc) = True Then
            DebeEntrar = True
            If Tipo = "DCL" Then
                'Doc = InputBox("Numero de Nota", Doc)
                strNumero = Format(Doc, "0000000")
            End If
            NumIng = strSuc & strPto & Format(Doc, "0000000")
        End If
    End If
'PEDIDO PARA RESERVAR MERCADERIA
'    If Tipo = "DCL" And ReservaNotaCredito = 1 Then
'        strSql = " Select COALESCE(max(ped_codigo)+1,'" & FormatoD0("1000010000001") & "') as num " & _
'                 " From pedido WITH (TABLOCKX) " & _
'                 " Where emp_codigo='" & strEmpresa & "' AND ped_codigo LIKE '" & FormatoD0(100001) & "%'" & _
'                 " GROUP BY emp_codigo"
'        clsCon_Def.Ejecutar (strSql), "M"
'        strPedido = clsCon_Def.adorec_Def("num")
'        strSql = " INSERT INTO pedido (emp_codigo, ped_codigo, per_codigo, ven_codigo,tar_cre_codigo,tar_cre_porcentaje, ped_fecha, " & _
'             " ped_estado, ped_subtotal, ped_observacion,cot_codigo,tipo_fac_codigo,ped_egr_bodega, ped_fechamod, ped_usumod) " & _
'             " VALUES ('" & strEmpresa & "'," & strPedido & ",'','', " & _
'             " '','', CURRENT_TIMESTAMP," & _
'             " '0',0,'" & NumIng & "NOTA CREDITO', '',1,'',CURRENT_TIMESTAMP, '" & strUsuario & "') "
'        clsCon_Def.Ejecutar (strSql), "M"
'    End If
    If DebeEntrar = True Then
        
        If strSuc = "" And strPto = "" Then
            strSuc = strSucursal
            strPto = strPtoFactura
            strNumero = Doc
        End If
        strDoc = NumIng
        If strNumero = "" Then strNumero = 0
        strSql = " INSERT INTO ingreso (ing_codigo,tip_ing_codigo,emp_codigo," & _
                 " ing_serie,ing_numero,ing_autorizacion,ing_caduca," & _
                 " for_pag_codigo,per_codigo,ing_fecha," & _
                 " ing_factura,ven_codigo,ing_subtotal,ing_subtotal_o,ing_dcto,ing_impuesto," & _
                 " ing_total,ing_observacion,ing_numasiento,ing_saldo,ing_fechamod,ing_usumod,ing_anulado,cod_iva_codigo) VALUES(" & _
                 " '" & strDoc & "','" & strTipo & "','" & strEmpresa & "'," & _
                 " '" & strSuc & strPto & "','" & strNumero & "','" & strAutor & "','" & strCaduca & "', " & _
                 " '" & ForPag & "','" & Persona & "','" & strFecha & "'," & _
                 " '" & Doc2 & "','" & Vendedor & "','" & Subtotal & "','" & SubTotal_0 & "','" & Dcto & "','" & Impto & "'," & _
                 " '" & Total & "','" & Observ & IIf(strIXC <> "", vbNewLine & "INGRESO POR CONTABILIZAR " & strIXC, "") & "','" & Asiento & "','" & saldo & "',CURRENT_TIMESTAMP,'" & strUsuario & "',0,'" & codIVA & "')"
        clsCon_Def.Ejecutar strSql, "M"
        dblSubTotal = Subtotal
        dblSubTotal0 = SubTotal_0
        dblDcto = Dcto
        dblIVA = Impto
        dblTotal = Total
        dblTotalProd = 0
        dblTotalServ = 0
        strPersona = Persona
        NuevoIng = True
    End If
    strSql = " COMMIT TRAN "
    clsCon_Def.Ejecutar strSql, "M"
    
End Function

Private Function NumeroDeNuevoDocIng(ConfirmaNumero As Boolean, strSuc As String, strPto As String, ByRef Doc As String) As Boolean
    Dim strSql As String
    Dim NumIng As Double
    Dim booPasar As Boolean
    Dim strDocume As String
    'LOCK TABLES ingreso WRITE,pedido WRITE,ingreso_salto i_s WRITE
    If strTipo = "DCL" Then
        strSql = " SELECT COALESCE(RIGHT(MAX(ing_codigo+1),7),1) AS num " & _
             " FROM ingreso WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND ing_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
             " AND ing_codigo NOT IN ( SELECT i_s.ing_codigo " & _
                                     " FROM ingreso_salto i_s " & _
                                     " WHERE i_s.emp_codigo='" & strEmpresa & "'" & _
                                     " AND i_s.tip_ing_codigo='" & strTipo & "' " & _
                                     " AND i_s.ing_codigo like '" & FormatoD0(strSuc & strPto) & "%')" & _
             " GROUP BY emp_codigo"
        'ConfirmaNumero = True
    ElseIf strSuc <> "" And strPto <> "" Then
        strSql = " SELECT COALESCE(RIGHT(MAX(ing_codigo+1),7),1) AS num " & _
             " FROM ingreso WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND ing_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
             " AND ing_codigo NOT IN ( SELECT i_s.ing_codigo " & _
                                     " FROM ingreso_salto i_s " & _
                                     " WHERE i_s.emp_codigo='" & strEmpresa & "'" & _
                                     " AND i_s.tip_ing_codigo='" & strTipo & "' " & _
                                     " AND i_s.ing_codigo like '" & FormatoD0(strSuc & strPto) & "%')" & _
             " GROUP BY emp_codigo"
    Else
        strSql = " SELECT COALESCE(RIGHT(MAX(ing_codigo+1),7),1) AS num " & _
             " FROM ingreso WITH (TABLOCKX) " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND ing_codigo like '" & strSuc & strPto & "%' " & _
             " AND ing_codigo NOT IN ( SELECT i_s.ing_codigo " & _
                                     " FROM ingreso_salto i_s " & _
                                     " WHERE i_s.emp_codigo='" & strEmpresa & "'" & _
                                     " AND i_s.tip_ing_codigo='" & strTipo & "' " & _
                                     " AND i_s.ing_codigo like '" & FormatoD0(strSuc & strPto) & "%')" & _
             " GROUP BY emp_codigo"
    End If
    clsCon_Def.Ejecutar strSql, "M"
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        Doc = FormatoD0(clsCon_Def.adorec_Def("num"))
    Else
        Doc = 1
    End If
    If ConfirmaNumero = True Then
        strDocume = NombreDoc
        booPasar = False
        NumeroDeNuevoDocIng = True
        While booPasar = False
            Doc = Val(InputBox(strDocume & " A GENERAR", "Numeración de Documento", Doc))
            If Doc = 0 Then
                 NumeroDeNuevoDocIng = False
                 Exit Function
            End If
            NumIng = strSuc & strPto & Format(Doc, "0000000")
            If NumIng = 0 Then
                booPasar = True
                NumeroDeNuevoDocIng = False
            Else
                strSql = " SELECT IIF(count(*) is null or count(*)=0,'N','S') as sn " & _
                         " FROM ingreso WITH (TABLOCKX) " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND tip_ing_codigo='" & strTipo & "' " & _
                         " AND ing_codigo='" & NumIng & "'"
                clsCon_Def.Ejecutar strSql, "M"
                If clsCon_Def.adorec_Def("sn") = "S" Then
                    If MsgBox("DOCUMENTO " & strDocume & " YA INGRESADO" & vbNewLine & "Compruebe la numeración." & _
                    vbNewLine & "SI, INGRESAR NUEVO NUMERO" & vbNewLine & _
                    "NO, PASAR Y NO INGRESAR", vbCritical + vbYesNo, "Numeración de Documento") = vbNo Then
                        booPasar = True
                        NumeroDeNuevoDocIng = False
                    End If
                Else
                    booPasar = True
                End If
            End If
        Wend
    Else
        NumeroDeNuevoDocIng = True
    End If
End Function

Public Sub NuevoDetIngRecargo(Recargo As String, Optional Precio As Double = 0)
    Dim strSql As String
    Recargo = Trim(Recargo)
    If Precio = 0 Then Exit Sub
    strSql = " SELECT det_ing_c_cantidad,det_ing_c_precio " & _
             " FROM det_ingreso_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & strDoc & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND oca_codigo='" & Recargo & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount = 0 Then
        strSql = " INSERT INTO det_ingreso_c (ing_codigo,tip_ing_codigo,emp_codigo,oca_codigo,det_ing_c_cantidad," & _
                 " det_ing_c_precio,det_ing_c_fechamod,det_ing_c_usumod) VALUES(" & _
                 " '" & strDoc & "','" & strTipo & "','" & strEmpresa & "','" & Recargo & "','1'," & _
                 " '" & Precio & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
    Else
        Dim cant As Double
        Dim Pre As Double
        cant = clsCon_Def.adorec_Def("det_ing_c_cantidad")
        Pre = clsCon_Def.adorec_Def("det_ing_c_precio")
        strSql = " UPDATE det_ingreso_c " & _
                 " SET det_ing_c_cantidad='1', " & _
                 " det_ing_c_precio='" & FormatoD4(cant * Pre + Precio) & "', " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ing_codigo='" & strDoc & "' " & _
                 " AND tip_ing_codigo='" & strTipo & "' " & _
                 " AND oca_codigo='" & Recargo & "' "
        clsCon_Def.Ejecutar strSql, "M"
    End If
End Sub

Public Sub NuevoDetEgrRecargo(Recargo As String, Optional Precio As Double = 0)
    Dim strSql As String
    Recargo = Trim(Recargo)
    If Precio = 0 Then Exit Sub
    strSql = " SELECT det_egr_c_cantidad,det_egr_c_precio " & _
             " FROM det_egreso_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & strDoc & "' " & _
             " AND tip_egr_codigo='" & strTipo & "' " & _
             " AND oca_codigo='" & Recargo & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount = 0 Then
        strSql = " INSERT INTO det_egreso_c (egr_codigo,tip_egr_codigo,emp_codigo,oca_codigo,det_egr_c_cantidad," & _
                 " det_egr_c_precio,det_egr_c_fechamod,det_egr_c_usumod) VALUES(" & _
                 " '" & strDoc & "','" & strTipo & "','" & strEmpresa & "','" & Recargo & "','1'," & _
                 " '" & Precio & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
    Else
        Dim cant As Double
        Dim Pre As Double
        cant = clsCon_Def.adorec_Def("det_egr_c_cantidad")
        Pre = clsCon_Def.adorec_Def("det_egr_c_precio")
        strSql = " UPDATE det_egreso_c " & _
                 " SET det_egr_c_cantidad='1', " & _
                 " det_egr_c_precio='" & FormatoD4(cant * Pre + Precio) & "', " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & strDoc & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND oca_codigo='" & Recargo & "' "
        clsCon_Def.Ejecutar strSql, "M"
    End If
End Sub

Public Sub NuevoDetIng(producto As String, Deposito As String, Cantidad As Double, Optional Precio As Double = 0, Optional Costo As Double = 0, Optional Dcto As Double = 0, Optional ConIva As Integer = 1)
    Dim strSql As String
    Dim clsContAux As New clsContenedor
    clsContAux.Inicializar AdoConn, AdoConnMaster
    producto = UCase(Trim(producto))
    Deposito = UCase(Trim(Deposito))
    If producto = "" Then Exit Sub
    If Deposito = "" Then Exit Sub
    If Cantidad = 0 Then Exit Sub
    If Precio = 0 Or Costo = 0 Then
        strSql = " SELECT prd_costo " & _
                 " FROM producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & producto & "'"
        clsCon_Def.Ejecutar strSql
        If Precio = 0 And strTipo <> "DCL" Then
            Precio = clsCon_Def.adorec_Def("prd_costo")
        End If
        If Costo = 0 Then
            Costo = clsCon_Def.adorec_Def("prd_costo")
        End If
    End If
    If Left(producto, 3) = "PR-" Then Costo = 0
    strSql = " SELECT det_ing_cantidad,det_ing_precio,det_ing_costo,det_ing_dcto " & _
             " FROM det_ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & strDoc & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND prd_codigo='" & producto & "' " & _
             " AND dep_codigo='" & Deposito & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount = 0 Then
        strSql = " INSERT INTO det_ingreso (emp_codigo,ing_codigo,tip_ing_codigo,prd_codigo,dep_codigo,det_ing_cantidad," & _
                 " det_ing_precio,det_ing_costo,det_ing_dcto,det_ing_fechamod,det_ing_usumod) VALUES(" & _
                 " '" & strEmpresa & "','" & strDoc & "','" & strTipo & "','" & producto & "','" & Deposito & "','" & Cantidad & "'," & _
                 " '" & Precio & "','" & Costo & "','" & Dcto & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
'PEDIDO PARA RESERVAR MERCADERIA
'        If strTipo = "DCL" And ReservaNotaCredito = 1 Then
'        strSql = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
'                 " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod) " & _
'                 " VALUES ('" & strEmpresa & "'," & strPedido & ",'" & producto & "','" & Deposito & "'," & cantidad & "," & cantidad & ", " & _
'                 " " & Precio & "," & Dcto & ", CURRENT_TIMESTAMP, '" & strUsuario & "') "
'        clsCon_Def.Ejecutar (strSql), "M"
'        End If
    Else
        Dim cant As Double
        Dim Pre As Double
        Dim Cos As Double
        Dim Des As Double
        cant = clsCon_Def.adorec_Def("det_ing_cantidad")
        Pre = clsCon_Def.adorec_Def("det_ing_precio")
        Cos = clsCon_Def.adorec_Def("det_ing_costo")
        Des = clsCon_Def.adorec_Def("det_ing_dcto")
        strSql = " UPDATE det_ingreso " & _
                 " SET det_ing_cantidad='" & cant + Cantidad & "', " & _
                 " det_ing_precio='" & FormatoD4((cant * Pre + Cantidad * Precio) / (cant + Cantidad)) & "', " & _
                 " det_ing_costo='" & FormatoD4((cant * Cos + Cantidad * Costo) / (cant + Cantidad)) & "', " & _
                 " det_ing_dcto='" & FormatoD4(Dcto + Des) & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ing_codigo='" & strDoc & "' " & _
                 " AND tip_ing_codigo='" & strTipo & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND dep_codigo='" & Deposito & "' "
        clsCon_Def.Ejecutar strSql, "M"
'PEDIDO PARA RESERVAR MERCADERIA
'        If strTipo = "DCL" And ReservaNotaCredito = 1 Then
'        strSql = " UPDATE det_pedido " & _
'                 " SET det_ped_cant_pedida=det_ped_cant_pedida+" & cantidad & "," & _
'                 " det_ped_cant_entregada=det_ped_cant_entregada+" & cantidad & "" & _
'                 " WHERE emp_codigo='" & strEmpresa & "' " & _
'                 " AND ped_codigo=" & strPedido & " " & _
'                 " AND prd_codigo='" & producto & "' " & _
'                 " AND dep_codigo='" & Deposito & "' "
'        clsCon_Def.Ejecutar (strSql), "M"
'        End If
    End If
    If Left(producto, 3) = "PR-" Then
        dblTotalServ = dblTotalServ + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        If ConIva = 1 Then
            dblTotalServIVA = dblTotalServIVA + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        End If
    Else
        If strIXC = "" Then
            If IngresaAutoAContenedor = True Then
                clsContAux.IngresarPrendas producto, Cantidad, Deposito, strTipo, strDoc
            End If
        Else
            clsContAux.CambiarDocumentoPrendas producto, "IXC", strIXC, strTipo, strDoc
        End If
        dblTotalProd = dblTotalProd + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        If ConIva = 1 Then
            dblTotalProdIVA = dblTotalProdIVA + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        End If
    End If
    If strTipo = "COM" Then
        strSql = " DELETE FROM persona_producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND per_codigo='" & strPersona & "'"
        clsCon_Def.Ejecutar strSql, "M"
        strSql = " INSERT INTO persona_producto (per_codigo,emp_codigo,prd_codigo,per_pro_precio,per_pro_fechamod,per_pro_usumod) " & _
                 " VALUES('" & strPersona & "','" & strEmpresa & "','" & producto & "','" & Precio & "',CURRENT_TIMESTAMP,'" & strUsuario & "') "
        clsCon_Def.Ejecutar strSql, "M"
    End If
    Existencia producto, Deposito, Cantidad
    clsCostear_Def.Recostear strFecha, , producto
End Sub
Public Function NombreDoc()
    Dim strSql As String
    Dim Tabla As String
    Dim Campo As String
    If strIE = "E" Then
        Tabla = "tipo_egreso"
        Campo = "tip_egr"
    Else
        Tabla = "tipo_ingreso"
        Campo = "tip_ing"
    End If
    strSql = " SELECT " & Campo & "_nombre FROM " & Tabla & _
             " WHERE " & Campo & "_codigo='" & strTipo & "' " & _
             " AND emp_codigo='" & strEmpresa & "'"
    clsCon_Def.Ejecutar strSql
    NombreDoc = clsCon_Def.adorec_Def(Campo & "_nombre")
End Function

Public Function NuevoEgr(IngAutoAContenedor As Boolean, Tipo As String, ConfirmaNumero As Boolean, Optional strSuc As String = "", Optional strPto As String = "", Optional strNumero As String = "", Optional ForPag As String = "", Optional Persona As String = "", Optional Fecha As String = "", Optional Doc2 As String = "", Optional Vendedor As String = "", Optional Observ As String = "", Optional Asiento As String = "", Optional strAutor As String = "", Optional strCaduca As String = "00/0000", Optional Subtotal As Double = 0, Optional SubTotal_0 As Double = 0, Optional Dcto As Double = 0, Optional Impto As Double = 0, Optional Total As Double = 0, Optional saldo As Double = 0, Optional SecPublico As Boolean = False, Optional SinIVA As Boolean = False, Optional codIVA As Integer = 0, Optional strFormaCobro As String = "") As Boolean
    Dim strSql As String
    Dim Doc As String
    Dim NumEgr As Double
    Dim strSucursalAux As String
    Dim strPtoAux As String
    strSucursalAux = strSuc
    strPtoAux = strPto
    booSecPublico = SecPublico
    booSinIva = SinIVA
    Observ = UCase(Observ)
    IngresaAutoAContenedor = IngAutoAContenedor
    strIE = "E"
    strTipo = Tipo
    If Fecha = "" Then
        strFecha = Format(HoyDia, "yyyy-mm-dd")
    Else
        strFecha = Fecha
    End If
    If strPto = "" Then
        strPto = strPtoFactura
    End If
    NuevoEgr = False
    strSql = " BEGIN TRAN"
    clsCon_Def.Ejecutar strSql, "M"
    If NumeroDeNuevoDocEgr(ConfirmaNumero, strSucursal, strPto, Doc) = True Then
        If Tipo = "ECA" Then
            Doc = strNumero
        End If
        If Tipo = "FAC" And ConfirmaNumero = False Then
            Doc = strNumero
        End If
        NumEgr = strSucursal & strPto & Format(Doc, "0000000")
        If strSuc = "" And strPto = "" Then
            strSuc = strSucursal
            strPto = strPtoFactura
            strNumero = Doc
        End If
        If Tipo = "DPV" Then
            strSuc = strSucursalAux
            strPto = strPtoAux
        End If
        If strNumero = "" Then strNumero = 0

        strDoc = NumEgr
        If booSinIva = True Then
            SubTotal_0 = Subtotal
            Subtotal = 0
            Total = SubTotal_0 - Dcto
        End If
        strSql = " INSERT INTO egreso (egr_codigo,tip_egr_codigo,emp_codigo," & _
                 " egr_serie,egr_numero,egr_autorizacion,egr_caduca," & _
                 " for_pag_codigo,per_codigo,egr_fecha," & _
                 " egr_factura,ven_codigo,egr_subtotal,egr_subtotal_o,egr_dcto,egr_impuesto," & _
                 " egr_total,egr_observacion,egr_numasiento,egr_saldo,egr_fechamod,egr_usumod,egr_anulado,cod_iva_codigo) VALUES(" & _
                 " '" & strDoc & "','" & strTipo & "','" & strEmpresa & "'," & _
                 " '" & strSuc & strPto & "','" & strNumero & "','" & strAutor & "','" & strCaduca & "', " & _
                 " '" & ForPag & "','" & Persona & "','" & strFecha & "'," & _
                 " '" & Doc2 & "','" & Vendedor & "','" & Subtotal & "','" & SubTotal_0 & "','" & Dcto & "','" & Impto & "'," & _
                 " '" & Total & "','" & Observ & "','" & Asiento & "','" & saldo & "',CURRENT_TIMESTAMP,'" & strUsuario & "',0,'" & codIVA & "')"
        clsCon_Def.Ejecutar strSql, "M"
        If strFormaCobro <> "" Then
            strSql = " INSERT INTO egreso_forma_cobro (emp_codigo,tip_egr_codigo,egr_codigo," & _
                     " for_cob_codigo,egr_for_cob_fechamod,egr_for_cob_usumod) VALUES(" & _
                     " '" & strEmpresa & "','" & strTipo & "','" & strDoc & "'," & _
                     " '" & strFormaCobro & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
            clsCon_Def.Ejecutar strSql, "M"
        End If
        dblSubTotal = Subtotal
        dblSubTotal0 = SubTotal_0
        dblDcto = Dcto
        dblIVA = Impto
        dblTotal = Total
        dblTotalProd = 0
        dblTotalServ = 0
        strPersona = Persona
        NuevoEgr = True
    End If
    strSql = " COMMIT TRAN"
    clsCon_Def.Ejecutar strSql, "M"
End Function

Public Sub DetRetenciones()
    Dim strSql As String
    Dim Tabla As String
    Dim ProdIVA As Double
    Dim ServIVA As Double
    If (dblTotalProdIVA + dblTotalServIVA) <> 0 Then
        ProdIVA = dblTotalProdIVA * dblIVA / (dblTotalProdIVA + dblTotalServIVA)
        ServIVA = dblTotalServIVA * dblIVA / (dblTotalProdIVA + dblTotalServIVA)
    Else
        ProdIVA = 0
        ServIVA = 0
    End If
    If strIE = "I" Then
        Tabla = "det_ingreso_ret"
    Else
        Tabla = "det_egreso_ret"
    End If
    strSql = " INSERT INTO " & Tabla & " " & _
             " SELECT retencion.emp_codigo,retencion.ret_codigo,'" & strDoc & "','" & strTipo & "', " & _
             " ROUND(retencion.ret_porcentaje * IIF(retencion.ret_gravara='SUBTOTAL'," & dblSubTotal & ",IIF(retencion.ret_gravara='SUBTOTALPRODUCTOS'," & dblTotalProd & ",IIF(retencion.ret_gravara='SUBTOTALSERVICIOS'," & dblTotalServ & ",IIF(retencion.ret_gravara='IVA'," & dblIVA & ",IIF(retencion.ret_gravara='IVAPRODUCTOS'," & ProdIVA & ",IIF(retencion.ret_gravara='IVASERVICIOS'," & ServIVA & ",IIF(retencion.ret_gravara='IVA0%'," & dblSubTotal0 & ",IIF(retencion.ret_gravara='Total'," & dblTotal & ",0))))))))/100.00,2),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
             " FROM retencion INNER JOIN persona_ret ON (retencion.ret_codigo = persona_ret.ret_codigo) " & _
             " AND (retencion.emp_codigo = persona_ret.emp_codigo) " & _
             " WHERE '" & strFecha & "' BETWEEN ret_fechaini AND ret_fechafin AND persona_ret.per_codigo='" & strPersona & "' AND persona_ret.emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSql, "M"
    If strIE = "I" Then
        strSql = " SELECT COALESCE(SUM(det_ing_ret_valor),0) as ret " & _
                 " FROM " & Tabla & " " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_ing_codigo='" & strTipo & "' " & _
                 " AND ing_codigo='" & strDoc & "' " & _
                 " GROUP BY emp_codigo "
    Else
        strSql = " SELECT COALESCE(SUM(det_egr_ret_valor),0) as ret " & _
                 " FROM " & Tabla & " " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND egr_codigo='" & strDoc & "' " & _
                 " GROUP BY emp_codigo "
    End If
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        dblRetencion = clsCon_Def.adorec_Def("ret")
    End If
End Sub

Public Sub ModificaEgr(Optional ForPag As String = "", Optional Persona As String = "", Optional Fecha As String = "", Optional Doc2 As String = "", Optional Vendedor As String = "", Optional Observ As String = "", Optional Asiento As String = "", Optional strAutor As String = "", Optional strCaduca As String = "", Optional CambiarValores As Boolean = False, Optional Subtotal As Double = 0, Optional SubTotal_0 As Double = 0, Optional Dcto As Double = 0, Optional Impto As Double = 0, Optional Total As Double = 0, Optional saldo As Double = 0)
    Dim strSql As String
    Dim Doc As String
    strSql = " UPDATE egreso SET "
    If ForPag <> "" Then
        strSql = strSql & " for_pag_codigo='" & ForPag & "',"
    End If
    If Persona <> "" Then
        strSql = strSql & " per_codigo='" & Persona & "',"
        strPersona = Persona
    End If
    If Fecha <> "" Then
        strSql = strSql & " egr_fecha='" & Fecha & "',"
    End If
    If Doc2 <> "" Then
        strSql = strSql & " egr_factura='" & Doc2 & "',"
    End If
    If Vendedor <> "" Then
        strSql = strSql & " ven_codigo='" & Vendedor & "',"
    End If
    If Observ <> "" Then
        strSql = strSql & " egr_observacion='" & Observ & "',"
    End If
    If Asiento <> "" Then
        strSql = strSql & " egr_numasiento='" & Asiento & "',"
    End If
    If strAutor <> "" Then
        strSql = strSql & " egr_autorizacion='" & strAutor & "',"
    End If
    If strCaduca <> "" Then
        strSql = strSql & " egr_caduca='" & strCaduca & "',"
    End If
    If CambiarValores = True Then
        strSql = strSql & " egr_subtotal='" & Subtotal & "',"
        strSql = strSql & " egr_subtotal_o='" & SubTotal_0 & "',"
        strSql = strSql & " egr_dcto='" & Dcto & "',"
        strSql = strSql & " egr_impuesto='" & Impto & "',"
        strSql = strSql & " egr_total='" & Total & "',"
        
        dblSubTotal = Subtotal
        dblSubTotal0 = SubTotal_0
        dblDcto = Dcto
        dblIVA = Impto
        dblTotal = Total
    End If
    If saldo <> 0 Then
            strSql = strSql & " egr_saldo='" & saldo & "',"
    End If
    strSql = Left(strSql, Len(strSql) - 1)
    strSql = strSql & " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='" & strTipo & "' " & _
             " AND egr_codigo='" & strDoc & "'"
    clsCon_Def.Ejecutar strSql, "M"
        
End Sub

Public Sub ModificaIng(Optional ForPag As String = "", Optional Persona As String = "", Optional Fecha As String = "", Optional Doc2 As String = "", Optional Vendedor As String = "", Optional Observ As String = "", Optional Asiento As String = "", Optional strAutor As String = "", Optional strCaduca As String = "", Optional CambiarValores As Boolean = False, Optional Subtotal As Double = 0, Optional SubTotal_0 As Double = 0, Optional Dcto As Double = 0, Optional Impto As Double = 0, Optional Total As Double = 0, Optional saldo As Double = 0)
    Dim strSql As String
    Dim Doc As String
    strSql = " UPDATE ingreso SET "
    If ForPag <> "" Then
        strSql = strSql & " for_pag_codigo='" & ForPag & "',"
    End If
    If Persona <> "" Then
        strSql = strSql & " per_codigo='" & Persona & "',"
        strPersona = Persona
    End If
    If Fecha <> "" Then
        strSql = strSql & " ing_fecha='" & Fecha & "',"
    End If
    If Doc2 <> "" Then
        strSql = strSql & " ing_factura='" & Doc2 & "',"
    End If
    If Vendedor <> "" Then
        strSql = strSql & " ven_codigo='" & Vendedor & "',"
    End If
    If Observ <> "" Then
        strSql = strSql & " ing_observacion='" & Observ & "',"
    End If
    If Asiento <> "" Then
        strSql = strSql & " ing_numasiento='" & Asiento & "',"
    End If
    If strAutor <> "" Then
        strSql = strSql & " ing_autorizacion='" & strAutor & "',"
    End If
    If strCaduca <> "" Then
        strSql = strSql & " ing_caduca='" & strCaduca & "',"
    End If
    If CambiarValores = True Then
        strSql = strSql & " ing_subtotal='" & Subtotal & "',"
        strSql = strSql & " ing_subtotal_o='" & SubTotal_0 & "',"
        strSql = strSql & " ing_dcto='" & Dcto & "',"
        strSql = strSql & " ing_impuesto='" & Impto & "',"
        strSql = strSql & " ing_total='" & Total & "',"
        
        dblSubTotal = Subtotal
        dblSubTotal0 = SubTotal_0
        dblDcto = Dcto
        dblIVA = Impto
        dblTotal = Total
    End If
    If saldo <> 0 Then
            strSql = strSql & " ing_saldo='" & saldo & "',"
    End If
    strSql = Left(strSql, Len(strSql) - 1)
    strSql = strSql & " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND ing_codigo='" & strDoc & "'"
    clsCon_Def.Ejecutar strSql, "M"
        
End Sub

Private Function NumeroDeNuevoDocEgr(ConfirmaNumero As Boolean, strSuc As String, strPto As String, ByRef Doc As String) As Boolean
    Dim strSql As String
    Dim NumEgr As Double
    Dim booPasar As Boolean
    Dim strDocume As String
    'LOCK TABLES egreso WRITE, factura_guia WRITE,egreso_forma_cobro WRITE
    If strTipo = "DPV" Then
        strPto = "001"
    End If
    If strTipo = "FAC" Then
        strSql = " SELECT COALESCE(RIGHT(MAX(egr_codigo+1),7)," & lngFacturaDesde & ") AS num " & _
                 " FROM egreso WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND egr_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
                 " AND egr_codigo BETWEEN '" & _
                 FormatoD0(strSuc & strPto & Format(lngFacturaDesde, "0000000")) & "' AND '" & _
                 FormatoD0(strSuc & strPto & Format(lngFacturaHasta, "0000000")) & "' " & _
                 " GROUP BY emp_codigo"
    ElseIf strTipo = "GRE" Then
        strSql = " SELECT COALESCE(RIGHT(MAX(egr_codigo+1),7),1) AS num " & _
                 " FROM egreso WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND egr_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
                 " GROUP BY emp_codigo"
        strSql = strSql & " UNION " & _
                 " SELECT COALESCE(RIGHT(MAX(fac_gui_codigo+1),7),1) AS num " & _
                 " FROM factura_guia WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND fac_gui_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
                 " GROUP BY emp_codigo"
        strSql = strSql & " ORDER BY num DESC "
    
    Else
        strSql = " SELECT COALESCE(RIGHT(MAX(egr_codigo+1),7),1) AS num " & _
                 " FROM egreso WITH (TABLOCKX) " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND egr_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
                 " GROUP BY emp_codigo"
    End If
    clsCon_Def.Ejecutar strSql, "M"
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        Doc = FormatoD0(clsCon_Def.adorec_Def("num"))
    Else
        If strTipo = "FAC" Then
            Doc = FormatoD0(Format(lngFacturaDesde, "0000000"))
        Else
            Doc = FormatoD0(strSuc & strPto & Format(1, "0000000"))
        End If
    End If
    If ConfirmaNumero = True Then
        strDocume = NombreDoc
        booPasar = False
        NumeroDeNuevoDocEgr = True
        While booPasar = False
            If strTipo = "FAC" Then
                If GeneraDocElec = 0 Then
                    Do
                        Doc = Val(InputBox(strDocume & " A GENERAR", "Numeración de Documento", Doc))
                    Loop While (Doc < lngFacturaDesde Or lngFacturaHasta < Doc) And Doc <> 0
                    If Doc = lngFacturaHasta Then
                        MsgBox "Se debió haber acabado el bloc de facturas." & vbNewLine & _
                        "Antes de continuar" & vbNewLine & "Ingrese el nuevo bloc de facturas", vbInformation, "Facturación"
                        IngresoBlocFactura
                    End If
                End If
            Else
                Doc = Val(InputBox(strDocume & " A GENERAR", "Numeración de Documento", Doc))
            End If
            
            NumEgr = strSuc & strPto & Format(Doc, "0000000")
            If NumEgr = strSuc & strPto & Format(0, "0000000") Then
                booPasar = True
                NumeroDeNuevoDocEgr = False
            Else
                strSql = " SELECT IIF(count(*) is null or count(*)=0,'N','S') as sn " & _
                         " FROM egreso WITH (TABLOCKX) " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND tip_egr_codigo='" & strTipo & "' " & _
                         " AND egr_codigo='" & NumEgr & "'"
                clsCon_Def.Ejecutar strSql, "M"
                If clsCon_Def.adorec_Def("sn") = "S" Then
                    If MsgBox("DOCUMENTO " & strDocume & " YA INGRESADO" & vbNewLine & "Compruebe la numeración." & _
                    vbNewLine & "SI, INGRESAR NUEVO NUMERO" & vbNewLine & _
                    "NO, PASAR Y NO INGRESAR", vbCritical + vbYesNo, "Numeración de Documento") = vbNo Then
                        booPasar = True
                        NumeroDeNuevoDocEgr = False
                    End If
                Else
                    booPasar = True
                End If
            End If
        Wend
    Else
        NumeroDeNuevoDocEgr = True
    End If
End Function

Public Sub EliminarDetEgr(producto As String, Deposito As String)
    Dim strSql As String
    producto = UCase(Trim(producto))
    Deposito = UCase(Trim(Deposito))
    If producto = "" Then Exit Sub
    If Deposito = "" Then Exit Sub
    strSql = " SELECT det_egr_cantidad,det_egr_precio,det_egr_costo,det_egr_dcto,prd_iva " & _
             " FROM det_egreso INNER JOIN producto ON det_egreso.emp_codigo=producto.emp_codigo " & _
             " AND det_egreso.prd_codigo=producto.prd_codigo " & _
             " WHERE det_egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & strDoc & "' " & _
             " AND tip_egr_codigo='" & strTipo & "' " & _
             " AND det_egreso.prd_codigo='" & producto & "' " & _
             " AND dep_codigo='" & Deposito & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount <> 0 Then
        Dim cant As Double
        Dim Pre As Double
        Dim Cos As Double
        Dim Des As Double
        Dim ConIva As Integer
        cant = clsCon_Def.adorec_Def("det_egr_cantidad")
        Pre = clsCon_Def.adorec_Def("det_egr_precio")
        Cos = clsCon_Def.adorec_Def("det_egr_costo")
        Des = clsCon_Def.adorec_Def("det_egr_dcto")
        ConIva = clsCon_Def.adorec_Def("prd_iva")
        strSql = " DELETE FROM det_egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & strDoc & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND dep_codigo='" & Deposito & "' "
        clsCon_Def.Ejecutar strSql, "M"
        If Left(producto, 3) = "PR-" Then
            dblTotalServ = dblTotalServ - (cant * Pre - Des)
            If ConIva = 1 Then
                dblTotalServIVA = dblTotalServIVA - (cant * Pre - Des)
            End If
        Else
            dblTotalProd = dblTotalProd - (cant * Pre - Des)
            If ConIva = 1 Then
                dblTotalProdIVA = dblTotalProdIVA - (cant * Pre - Des)
            End If
        End If
        Existencia producto, Deposito, (cant)
        clsCostear_Def.Recostear strFecha, , producto
    End If
End Sub
Public Sub EliminarDetIng(producto As String, Deposito As String)
    Dim strSql As String
    producto = UCase(Trim(producto))
    Deposito = UCase(Trim(Deposito))
    If producto = "" Then Exit Sub
    If Deposito = "" Then Exit Sub
    strSql = " SELECT det_ing_cantidad,det_ing_precio,det_ing_costo,det_ing_dcto,prd_iva " & _
             " FROM det_ingreso INNER JOIN producto ON det_ingreso.emp_codigo=producto.emp_codigo " & _
             " AND det_ingreso.prd_codigo=producto.prd_codigo " & _
             " WHERE det_ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & strDoc & "' " & _
             " AND tip_ing_codigo='" & strTipo & "' " & _
             " AND det_ingreso.prd_codigo='" & producto & "' " & _
             " AND dep_codigo='" & Deposito & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount <> 0 Then
        Dim cant As Double
        Dim Pre As Double
        Dim Cos As Double
        Dim Des As Double
        Dim ConIva As Integer
        cant = clsCon_Def.adorec_Def("det_ing_cantidad")
        Pre = clsCon_Def.adorec_Def("det_ing_precio")
        Cos = clsCon_Def.adorec_Def("det_ing_costo")
        Des = clsCon_Def.adorec_Def("det_ing_dcto")
        ConIva = clsCon_Def.adorec_Def("prd_iva")
        strSql = " DELETE FROM det_ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ing_codigo='" & strDoc & "' " & _
                 " AND tip_ing_codigo='" & strTipo & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND dep_codigo='" & Deposito & "' "
        clsCon_Def.Ejecutar strSql, "M"
        If Left(producto, 3) = "PR-" Then
            dblTotalServ = dblTotalServ - (cant * Pre * -Des)
            If ConIva = 1 Then
                dblTotalServIVA = dblTotalServIVA - (cant * Pre * -Des)
            End If
        Else
            dblTotalProd = dblTotalProd - (cant * Pre - Des)
            If ConIva = 1 Then
                dblTotalProdIVA = dblTotalProdIVA - (cant * Pre * -Des)
            End If
        End If
        Existencia producto, Deposito, (cant * -1)
        clsCostear_Def.Recostear strFecha, , producto
    End If
End Sub

Public Sub NuevoDetEgr(producto As String, Deposito As String, Cantidad As Double, Optional Precio As Double = 0, Optional Costo As Double = 0, Optional Dcto As Double = 0, Optional ConIva As Integer = 1, Optional Pdcto As String = "NULL")
    Dim strSql As String
    Dim clsContAux As New clsContenedor
    clsContAux.Inicializar AdoConn, AdoConnMaster
    producto = UCase(Trim(producto))
    Deposito = UCase(Trim(Deposito))
    If producto = "" Then Exit Sub
    If Deposito = "" Then Exit Sub
    If Cantidad = 0 Then Exit Sub
    If Precio = 0 Or Costo = 0 Then
        strSql = " SELECT prd_costo " & _
                 " FROM producto " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & producto & "'"
        clsCon_Def.Ejecutar strSql
        If Precio = 0 Then
            If strTipo <> "FAC" Then
                Precio = clsCon_Def.adorec_Def("prd_costo")
            End If
        End If
        If Costo = 0 Then
            Costo = clsCon_Def.adorec_Def("prd_costo")
        End If
    End If
    If Pdcto <> "NULL" Then
        Pdcto = "'" & Pdcto & "'"
    End If
    If Left(producto, 3) = "PR-" Then Costo = 0
    strSql = " SELECT det_egr_cantidad,det_egr_precio,det_egr_costo,det_egr_dcto " & _
             " FROM det_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & strDoc & "' " & _
             " AND tip_egr_codigo='" & strTipo & "' " & _
             " AND prd_codigo='" & producto & "' " & _
             " AND dep_codigo='" & Deposito & "' "
    clsCon_Def.Ejecutar strSql
    If clsCon_Def.adorec_Def.RecordCount = 0 Then
        If FormatoD2(Replace(Pdcto, "'", "")) > 100 Then
            'Precio = 0.01 ' InputBox("Producto:" & producto & " - Egreso:" & strDoc & " - Precio:" & Precio & " - Dcto:" & Dcto, "Precio")
            Pdcto = 100 'InputBox("Producto:" & producto & " - Egreso:" & strDoc & " - Precio:" & Precio & " - Dcto:" & Dcto, "%DCTO")
        End If
        strSql = " INSERT INTO det_egreso (emp_codigo,egr_codigo,tip_egr_codigo,prd_codigo,dep_codigo,det_egr_cantidad," & _
                 " det_egr_precio,det_egr_costo,det_egr_dcto,det_egr_pdcto,det_egr_fechamod,det_egr_usumod) VALUES(" & _
                 " '" & strEmpresa & "','" & strDoc & "','" & strTipo & "','" & producto & "','" & Deposito & "','" & Cantidad & "'," & _
                 " '" & Precio & "','" & Costo & "','" & Dcto & "'," & Pdcto & ",CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSql, "M"
    Else
        Dim cant As Double
        Dim Pre As Double
        Dim Cos As Double
        Dim Des As Double
        cant = clsCon_Def.adorec_Def("det_egr_cantidad")
        Pre = clsCon_Def.adorec_Def("det_egr_precio")
        Cos = clsCon_Def.adorec_Def("det_egr_costo")
        Des = clsCon_Def.adorec_Def("det_egr_dcto")
        strSql = " UPDATE det_egreso " & _
                 " SET det_egr_cantidad='" & cant + Cantidad & "', " & _
                 " det_egr_precio='" & FormatoD4((cant * Pre + Cantidad * Precio) / (cant + Cantidad)) & "', " & _
                 " det_egr_costo='" & FormatoD4(((cant * Cos + Cantidad * Costo) / (cant + Cantidad))) & "', " & _
                 " det_egr_dcto='" & FormatoD4(Dcto + Des) & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & strDoc & "' " & _
                 " AND tip_egr_codigo='" & strTipo & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND dep_codigo='" & Deposito & "' "
        clsCon_Def.Ejecutar strSql, "M"
    End If
    If Left(producto, 3) = "PR-" Then
        dblTotalServ = dblTotalServ + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        If ConIva = 1 Then
            dblTotalServIVA = dblTotalServIVA + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        End If
    
    Else
        If IngresaAutoAContenedor = True Then
            'clsContAux.EgresarPrendas producto, cantidad, "PRI','DE", strTipo, strDoc
            clsContAux.EgresarPrendas producto, Cantidad, Deposito, strTipo, strDoc
        End If
        dblTotalProd = dblTotalProd + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        If ConIva = 1 Then
            dblTotalProdIVA = dblTotalProdIVA + FormatoD2(FormatoD4(Cantidad) * FormatoD8(Precio)) - FormatoD4(Dcto)
        End If
    End If
    Existencia producto, Deposito, (Cantidad * -1)
    If strTipo <> "FAC" Then
        clsCostear_Def.Recostear strFecha, , producto
    End If
End Sub

Private Sub Existencia(producto As String, Deposito As String, Cantidad As Double)
    Dim clsCon_DefAUX As New clsConsulta
    Dim strSql As String
    Dim disponible As Double
    Dim CantFinal As Double
    producto = UCase(Trim(producto))
    Deposito = UCase(Trim(Deposito))
    If producto = "" Then Exit Sub
    If Deposito = "" Then Exit Sub
    If Cantidad = 0 Then Exit Sub
    
    strSql = " SELECT exi_cantidad " & _
             " FROM existencia " & _
             " WHERE prd_codigo='" & producto & "' " & _
             " AND dep_codigo='" & Deposito & "' " & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsCon_DefAUX.Inicializar AdoConn, AdoConnMaster
    clsCon_DefAUX.Ejecutar strSql
    If clsCon_DefAUX.adorec_Def.RecordCount > 0 Then
        While (Not clsCon_DefAUX.adorec_Def.EOF) And Cantidad <> 0
            disponible = clsCon_DefAUX.adorec_Def("exi_cantidad")
            
            If disponible < 0 Then
            'si es existencia negativa
                If disponible + Cantidad >= 0 Then
                'para poner la existencia negativa en cero y continuar
                    CantFinal = 0
                    Cantidad = Cantidad + disponible
                ElseIf disponible + Cantidad < 0 Then
                'para poner la existencia negativa mas cerca de cero
                    CantFinal = disponible + Cantidad
                    Cantidad = 0
                End If
            ElseIf disponible > 0 Then
            'si existencia es positia
                If disponible + Cantidad >= 0 Then
                'si quedara mas de 0 existencias
                    CantFinal = disponible + Cantidad
                    Cantidad = 0
                ElseIf disponible + Cantidad < 0 Then
                'si quedara en 0 existencias y continua
                    CantFinal = 0
                    Cantidad = Cantidad + disponible
                End If
            Else
                CantFinal = Cantidad
                Cantidad = 0
            End If
            strSql = " UPDATE existencia " & _
                     " SET exi_cantidad='" & CantFinal & "', " & _
                     " exi_fechamod=CURRENT_TIMESTAMP, " & _
                     " exi_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & producto & "' " & _
                     " AND dep_codigo='" & Deposito & "' "
            clsCon_Def.Ejecutar strSql, "M"
            clsCon_DefAUX.adorec_Def.MoveNext
        Wend
    'INSERT
    End If
    If Cantidad <> 0 Then
        strSql = " SELECT exi_cantidad " & _
                 " FROM existencia " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & producto & "' " & _
                 " AND dep_codigo='" & Deposito & "'"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def.RecordCount <= 0 Then
            strSql = " INSERT INTO existencia(prd_codigo,dep_codigo,emp_codigo," & _
                     " exi_cantidad,exi_fechamod,exi_usumod) VALUES(" & _
                     " '" & producto & "','" & Deposito & "','" & strEmpresa & "', " & _
                     " '" & Cantidad & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        Else
            strSql = " UPDATE existencia " & _
                     " SET exi_cantidad=exi_cantidad+'" & Cantidad & "', " & _
                     " exi_fechamod=CURRENT_TIMESTAMP, " & _
                     " exi_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND prd_codigo='" & producto & "' " & _
                     " AND dep_codigo='" & Deposito & "' "
        End If
        clsCon_Def.Ejecutar strSql, "M"
        Cantidad = 0
    End If
    Set clsCon_DefAUX = Nothing
End Sub

Public Function AnularIng(Optional NumIng As String = "", Optional TipIng As String = "", Optional NumAsiento As String = "", Optional Motivo As String = "") As Boolean
    Dim strSql As String
    Dim strFecha As String
    Dim Anul As Boolean
    Dim clsAsiento As New clsContable
    Anul = False
    strDoc = NumIng
    strTipo = TipIng
    If NumIng <> "" And TipIng <> "" Then
        strSql = " SELECT COALESCE(ing_numasiento,'') as ing_numasiento,ing_fecha " & _
                 " FROM ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ing_codigo='" & NumIng & "'" & _
                 " AND tip_ing_codigo='" & TipIng & "'"
        clsCon_Def.Ejecutar strSql
        NumAsiento = clsCon_Def.adorec_Def("ing_numasiento")
        strFecha = clsCon_Def.adorec_Def("ing_fecha")
    Else
        strSql = " SELECT COALESCE(ing_codigo,'') as ing_codigo,COALESCE(tip_ing_codigo,'') as tip_ing_codigo,ing_fecha " & _
                 " FROM ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ing_numasiento='" & NumAsiento & "'"
        clsCon_Def.Ejecutar strSql
        NumIng = clsCon_Def.adorec_Def("ing_codigo")
        TipIng = clsCon_Def.adorec_Def("tip_ing_codigo")
        strFecha = clsCon_Def.adorec_Def("ing_fecha")
    End If
    If NumAsiento <> "" Then
        strSql = " SELECT COALESCE(COUNT(*),0) as num,asi_numasiento " & _
                 " FROM asiento " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND asi_numasiento='" & NumAsiento & "'" & _
                 " AND asi_descripcion LIKE 'ASIENTO ANULADO%' GROUP BY asi_numasiento"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def("num") = 0 Then
            If MsgBox("El asiento " & NumAsiento & " aun no esta anulado.", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
                clsAsiento.Inicializar AdoConn, AdoConnMaster
                Motivo = ""
                While Motivo = ""
                    Motivo = InputBox("Motivo de Anulacion", "Contabilidad")
                    If Motivo = "" Then
                        If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el Asiento?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                            Anul = False
                            Motivo = "NO ANULAR"
                        End If
                    Else
                        Anul = True
                    End If
                Wend
                If Anul = True Then
                    clsAsiento.NumAsiento = NumAsiento
                    clsAsiento.AnularAsientoYOtros UCase(Motivo), clsCon_Def.adorec_Def("asi_descripcion")
                End If
            End If
        Else
            Anul = True
        End If
    
    Else
        Anul = True
    End If
    If Anul = True Then
        Dim clsCon_Aux As New clsConsulta
        strSql = " UPDATE ingreso " & _
                 " SET ing_anulado=1,ing_observacion=CONCAT('" & Motivo & vbNewLine & "',ing_observacion)" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ing_codigo='" & NumIng & "'" & _
                 " AND tip_ing_codigo='" & TipIng & "'"
        clsCon_Def.Ejecutar strSql, "M"
        clsCon_Aux.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT prd_codigo,dep_codigo,det_ing_cantidad " & _
                 " FROM det_ingreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND ing_codigo='" & NumIng & "'" & _
                 " AND tip_ing_codigo='" & TipIng & "'"
        clsCon_Aux.Ejecutar strSql
        While Not clsCon_Aux.adorec_Def.EOF
            Existencia clsCon_Aux.adorec_Def("prd_codigo"), clsCon_Aux.adorec_Def("dep_codigo"), (clsCon_Aux.adorec_Def("det_ing_cantidad") * -1)
            clsCostear_Def.Recostear strFecha, , clsCon_Aux.adorec_Def("prd_codigo")
            clsCon_Aux.adorec_Def.MoveNext
        Wend
        
        Dim clsCont As New clsContenedor
        clsCont.Inicializar AdoConn, AdoConnMaster
        clsCont.AnularMovimiento TipIng, NumIng
        Set clsCont = Nothing
        
        Set clsCon_Aux = Nothing
        AnularIng = True
    Else
        AnularIng = False
    End If
End Function

Public Function AnularEgr(Optional NumEgr As String = "", Optional TipEgr As String = "", Optional NumAsiento As String = "", Optional Motivo As String = "") As Boolean
    Dim strSql As String
    Dim strFecha As String
    Dim Anul As Boolean
    Dim clsAsiento As New clsContable
    strDoc = NumEgr
    strTipo = TipEgr
    Anul = False
    If NumEgr <> "" And TipEgr <> "" Then
        strSql = " SELECT COALESCE(egr_numasiento,'') as egr_numasiento,egr_fecha " & _
                 " FROM egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND egr_codigo='" & NumEgr & "'" & _
                 " AND tip_egr_codigo='" & TipEgr & "'"
        clsCon_Def.Ejecutar strSql
        NumAsiento = clsCon_Def.adorec_Def("egr_numasiento")
        strFecha = clsCon_Def.adorec_Def("egr_fecha")
    Else
        strSql = " SELECT COALESCE(egr_codigo,'') as egr_codigo,COALESCE(tip_egr_codigo,'') as tip_egr_codigo,egr_fecha " & _
                 " FROM egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND egr_numasiento='" & NumAsiento & "'"
        clsCon_Def.Ejecutar strSql
        NumEgr = clsCon_Def.adorec_Def("egr_codigo")
        TipEgr = clsCon_Def.adorec_Def("tip_egr_codigo")
        strFecha = clsCon_Def.adorec_Def("egr_fecha")
    End If
    If NumAsiento <> "" Then
        strSql = " SELECT COALESCE(COUNT(*),0) as num,asi_numasiento " & _
                 " FROM asiento " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND asi_numasiento='" & NumAsiento & "'" & _
                 " AND asi_descripcion LIKE 'ASIENTO ANULADO%' GROUP BY asi_numasiento"
        clsCon_Def.Ejecutar strSql
        If clsCon_Def.adorec_Def("num") = 0 Then
            If MsgBox("El asiento " & NumAsiento & " aun no esta anulado.", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
                clsAsiento.Inicializar AdoConn, AdoConnMaster
                Motivo = ""
                While Motivo = ""
                    Motivo = InputBox("Motivo de Anulacion", "Contabilidad")
                    If Motivo = "" Then
                        If MsgBox("Debe ingresar un motivo para realizar la Anulación" & vbNewLine & "Desea Anular el Asiento?", vbQuestion + vbYesNo, "Contabilidad") = vbNo Then
                            Anul = False
                            Motivo = "NO ANULAR"
                        End If
                    Else
                        Anul = True
                    End If
                Wend
                If Anul = True Then
                    clsAsiento.NumAsiento = NumAsiento
                    clsAsiento.AnularAsientoYOtros Motivo, clsCon_Def.adorec_Def("asi_descripcion")
                    Anul = False
                End If
            End If
        Else
            Anul = True
        End If
    Else
        Anul = True
    End If
    If Anul = True Then
        Dim clsCon_Aux As New clsConsulta
        strSql = " UPDATE egreso " & _
                 " SET egr_anulado=1 ,egr_observacion=CONCAT('" & Motivo & vbNewLine & "',egr_observacion)" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND egr_codigo='" & NumEgr & "'" & _
                 " AND tip_egr_codigo='" & TipEgr & "'"
        clsCon_Def.Ejecutar strSql, "M"
        clsCon_Aux.Inicializar AdoConn, AdoConnMaster
        strSql = " SELECT prd_codigo,dep_codigo,det_egr_cantidad " & _
                 " FROM det_egreso " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND egr_codigo='" & NumEgr & "'" & _
                 " AND tip_egr_codigo='" & TipEgr & "'"
        clsCon_Aux.Ejecutar strSql
        While Not clsCon_Aux.adorec_Def.EOF
            Existencia clsCon_Aux.adorec_Def("prd_codigo"), clsCon_Aux.adorec_Def("dep_codigo"), (clsCon_Aux.adorec_Def("det_egr_cantidad"))
            clsCostear_Def.Recostear strFecha, , clsCon_Aux.adorec_Def("prd_codigo")
            clsCon_Aux.adorec_Def.MoveNext
        Wend
        Dim clsCont As New clsContenedor
        clsCont.Inicializar AdoConn, AdoConnMaster
        clsCont.AnularMovimiento TipEgr, NumEgr
        Set clsCont = Nothing
        Set clsCon_Aux = Nothing
        AnularEgr = True
    Else
        AnularEgr = False
    End If
End Function

Public Sub RevisarProductosEnPedidosReprogramados()
    Dim strSql As String
    Dim codEgr As String
    Dim booGuardar As Boolean
    Dim clsAux As New clsConsulta
    Dim clsEgresoAUX As New clsInventario
    Dim clsIngresoAux As New clsInventario
    booGuardar = False
    strSql = " SELECT repro.prd_codigo,(SUM(repro.cant)-exi_cantidad) as cant" & _
             " FROM (SELECT prd_codigo,SUM(det_ped_cant_programada) as cant " & _
             " FROM pedido INNER JOIN det_pedido " & _
             " ON pedido.emp_codigo=det_pedido.emp_codigo" & _
             " AND pedido.ped_codigo=det_pedido.ped_codigo" & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
             " AND pedido.ped_estado in (0,1)" & _
             " AND det_pedido.prd_codigo in (" & _
             " SELECT DISTINCT prd_codigo FROM det_ingreso WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='" & strTipo & "' AND ing_codigo='" & strDoc & "'" & _
             ")" & _
             " GROUP BY prd_codigo " & _
             " UNION " & _
             " SELECT prd_codigo,SUM(det_ped_cant_pedida) " & _
             " FROM pedido INNER JOIN det_pedido " & _
             " ON pedido.emp_codigo=det_pedido.emp_codigo" & _
             " AND pedido.ped_codigo=det_pedido.ped_codigo" & _
             " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
             " AND pedido.ped_estado in (-4)" & _
             " AND det_pedido.prd_codigo in (" & _
             " SELECT DISTINCT prd_codigo FROM det_ingreso WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='" & strTipo & "' AND ing_codigo='" & strDoc & "'" & _
             ")" & _
             " GROUP BY prd_codigo) repro INNER JOIN existencia ON existencia.emp_codigo='" & strEmpresa & "'" & _
             " AND repro.prd_codigo=existencia.prd_codigo AND existencia.dep_codigo='VEF'" & _
             " GROUP BY prd_codigo,exi_cantidad " & _
             " HAVING (SUM(repro.cant)-exi_cantidad)>0 "
    clsCon_Def.Ejecutar strSql
    
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        clsAux.Inicializar AdoConn, AdoConnMaster
        clsEgresoAUX.Inicializar AdoConn, AdoConnMaster
        clsIngresoAux.Inicializar AdoConn, AdoConnMaster
        
        booGuardar = clsEgresoAUX.NuevoEgr(True, "ETR", False, strSucursal, strPtoFactura, , , , Format(strFecha, "yyyy-MM-dd"), , , "RESERVA MERCADERIA REPROGRAMADA")
        booGuardar = clsIngresoAux.NuevoIng(True, "ITR", False, strSucursal, strPtoFactura, , , , Format(strFecha, "yyyy-MM-dd"), , , "RESERVA MERCADERIA REPROGRAMADA")
        If booGuardar Then
            codEgr = clsEgresoAUX.strDoc
            strContenedorRecurrente = "111"
            While Not clsCon_Def.adorec_Def.EOF
                strSql = " SELECT dep_codigo,prd_codigo,IIF(det_ing_cantidad<" & (FormatoD4(clsCon_Def.adorec_Def("cant"))) & "," & _
                         " det_ing_cantidad," & (FormatoD4(clsCon_Def.adorec_Def("cant"))) & ") as cant,det_ing_precio " & _
                         " FROM det_ingreso " & _
                         " WHERE emp_codigo='" & strEmpresa & "' AND tip_ing_codigo='" & strTipo & "' " & _
                         " AND ing_codigo='" & strDoc & "' AND prd_codigo='" & clsCon_Def.adorec_Def("prd_codigo") & "'"
                clsAux.Ejecutar strSql
                clsEgresoAUX.NuevoDetEgr clsAux.adorec_Def("prd_codigo"), clsAux.adorec_Def("dep_codigo"), FormatoD4(clsAux.adorec_Def("cant")), FormatoD4(clsAux.adorec_Def("det_ing_precio")), FormatoD4(clsAux.adorec_Def("det_ing_precio"))
                clsIngresoAux.NuevoDetIng clsAux.adorec_Def("prd_codigo"), "VEF", FormatoD4(clsAux.adorec_Def("cant")), FormatoD4(clsAux.adorec_Def("det_ing_precio")), FormatoD4(clsAux.adorec_Def("det_ing_precio"))
            Wend
            InicializarContenedorRecurrente
        End If
    
        Set clsEgresoAUX = Nothing
        Set clsIngresoAux = Nothing
        
        Dim frmEntrega As New frmReporte
        frmEntrega.strNumero = codEgr
        frmEntrega.strTipo = "ETR"
        frmEntrega.strReporte = "rptTransferencia"
        frmEntrega.Show
    End If
    InicializarContenedorRecurrente
    
End Sub

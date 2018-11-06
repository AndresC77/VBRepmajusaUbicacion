VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtaXx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsCon_Def As clsConsulta
Public strTipoCta As String
Public strNoCta As String
Public strTipoDoc As String
Public strSustento As String
Public strFechaEmision As String
Public strFechaVence As String
Public strPersona As String
Public dblCero As Double
Public dblcodIVA As Double
Public dblIVA As Double
Public dblBaseICE As Double
Public dblcodICE As Double
Public dblICE As Double
Public dblTotal As Double
Public dblTotalProd As Double
Public dblTotalServ As Double
Public dblTotalProdIVA As Double
Public dblTotalServIVA As Double
Private clsCostear_Def As clsCostear

Public Sub AplicaNC2(strDoc As String, codCXC As String, ValAplicado As Double)
    Dim strSQL As String
    Dim SaldoCta As Double
    Dim SaldoNC As Double
    Dim FechaNC As String
    Dim NumNC As String
    Dim ClienteNC As String
    Dim AsientoNC As String
    Dim TotalNC As Double
    Dim maxpag As Integer
    Dim ValorPago As Double
    Dim Desc As String
    Dim clsCod As New clsConsulta
    Dim clsPag As New clsConsulta
    Dim clsInventa As New clsInventario
    clsCod.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsInventa.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT CONCAT(per_apellido,' ', per_nombre,' (',per_ruc,')') as cli " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C'" & _
             " AND per_codigo='" & strPersona & "'"
    clsCon_Def.Ejecutar strSQL
    ClienteNC = clsCon_Def.adorec_Def("cli")
        
    strSQL = " SELECT (ing_total-ing_saldo) as saldo,ing_total as total,CONCAT(ing_serie,'-',ing_numero) as nc,ing_fecha,ing_numasiento " & _
             " FROM  ingreso " & _
             " WHERE per_codigo IN ('" & strPersona & "') AND emp_codigo = '" & strEmpresa & "' " & _
             " AND tip_ing_codigo = 'DCL' AND ing_codigo='" & strDoc & "' "
    clsCon_Def.Ejecutar strSQL
    SaldoNC = FormatoD2(clsCon_Def.adorec_Def("saldo"))
    TotalNC = FormatoD2(clsCon_Def.adorec_Def("total"))
    FechaNC = clsCon_Def.adorec_Def("ing_fecha")
    NumNC = clsCon_Def.adorec_Def("nc")
    AsientoNC = clsCon_Def.adorec_Def("ing_numasiento")
        If SaldoNC > 0 Then
            strSQL = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d,tip_doc_cue_descripcion " & _
                     " FROM cuenta_p_c INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
                     " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo" & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " WHERE cuenta_p_c.cue_p_c_codigo = " & codCXC & " AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,com_ret_total,tip_doc_cue_descripcion,cuenta_p_c.cue_p_c_tipo HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)!=0 " & _
                     " ORDER BY cuenta_p_c.cue_p_c_fechaemision DESC,cue_p_c_egr_codigo DESC,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
                SaldoCta = FormatoD2(clsCon_Def.adorec_Def("d"))
                Desc = UCase(clsCon_Def.adorec_Def("tip_doc_cue_descripcion")) & ": " & clsCon_Def.adorec_Def("cue_p_c_egr_codigo")
                'Calcula el máximo codigo de pago para la cuenta
                strSQL = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                         " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                         " AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                         " AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                         " WHERE cuenta_p_c.cue_p_c_codigo= '" & codCXC & "' AND pago.emp_codigo = '" & strEmpresa & "' " & _
                         " AND pago.cue_p_c_tipo = 'C'" & _
                         " GROUP BY pago.emp_codigo"
                clsCod.Ejecutar strSQL
                If clsCod.adorec_Def.EOF Then
                    maxpag = 1
                Else
                    maxpag = clsCod.adorec_Def("pag") + 1
                End If
                    
                If SaldoNC >= SaldoCta Then
                    ValorPago = FormatoD2(ValAplicado)
                    Desc = Desc & " CANCELA"
                Else
                    ValorPago = FormatoD2(ValAplicado)
                    Desc = Desc & " ABONA"
                End If
                
                
                strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, " & _
                         " pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "', '" & Val(codCXC) & "', 'C', '" & Val(maxpag) & "', '" & FechaNC & "', '" & ValorPago & "', " & _
                         " '" & NumNC & "', '" & UCase(ClienteNC & " - " & Desc & " - NOTA DE CRÉDITO " & NumNC & " (" & strDoc & ")") & "', " & _
                         " '','" & AsientoNC & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsPag.Ejecutar strSQL, "M"
                
                
                strSQL = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                         " FROM cuenta_p_c INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
                         " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo" & _
                         " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                         " WHERE per_codigo IN ('" & strPersona & "') AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                         " AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
                         " AND cuenta_p_c.cue_p_c_codigo='" & Val(codCXC) & "'" & _
                         " GROUP BY cuenta_p_c.cue_p_c_codigo, cue_p_c_valor, com_ret_total "
                clsPag.Ejecutar strSQL, "M"
                
                If FormatoD2(clsPag.adorec_Def("d")) <= 0 Then
                    strSQL = " SELECT MAX(pag_fecha) as fec " & _
                             " FROM pago " & _
                             " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & Val(codCXC) & "' " & _
                             " AND emp_codigo = '" & strEmpresa & "' " & _
                             " AND pag_observacion NOT LIKE 'ANULADO%' " & _
                             " GROUP BY emp_codigo"
                    clsPag.Ejecutar strSQL, "M"
                    strSQL = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_fechapago='" & clsPag.adorec_Def("fec") & "', " & _
                             " cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, " & _
                             " cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & Val(codCXC) & "' " & _
                             " AND emp_codigo = '" & strEmpresa & "' "
                    clsPag.Ejecutar strSQL, "M"
                
                End If
                
                clsInventa.strTipo = "DCL"
                clsInventa.strIE = "I"
                clsInventa.strDoc = strDoc
                clsInventa.ModificaIng , , , , , , , , , , , , , , , FormatoD2(TotalNC - SaldoNC + ValAplicado)
            End If
        End If
End Sub

Public Sub AplicaNC(strDoc As String, Optional FechaDesde As String = "", Optional FechaHasta As String = "", Optional Fact As String = "")
    Dim strSQL As String
    Dim SaldoNC As Double
    Dim TotalNC As Double
    Dim Cta As String
    Dim FechaNC As String
    Dim NumNC As String
    Dim ClienteNC As String
    Dim AsientoNC As String
    Dim SaldoCta As Double
    Dim maxpag As Integer
    Dim ValorPago As Double
    Dim Desc As String
    Dim Filtro As String
    Dim clsCod As New clsConsulta
    Dim clsPag As New clsConsulta
    Dim clsInventa As New clsInventario
    clsCod.Inicializar AdoConn, AdoConnMaster
    clsPag.Inicializar AdoConn, AdoConnMaster
    clsInventa.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT per_aplica_nc,CONCAT(per_apellido,' ', per_nombre,' (',per_ruc,')') as cli " & _
             " FROM persona " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND cat_p_tipo='C'" & _
             " AND per_codigo='" & strPersona & "'"
    clsCon_Def.Ejecutar strSQL
    ClienteNC = clsCon_Def.adorec_Def("cli")
    If Abs(FormatoD0(clsCon_Def.adorec_Def("per_aplica_nc"))) <> 0 Or Fact <> "" Then
        
        strSQL = " SELECT (ing_total-ing_saldo) as saldo,ing_total as total,CONCAT(ing_serie,'-',ing_numero) as nc,ing_fecha,ing_numasiento " & _
                 " FROM  ingreso " & _
                 " WHERE per_codigo IN ('" & strPersona & "') AND emp_codigo = '" & strEmpresa & "' " & _
                 " AND tip_ing_codigo = 'DCL' AND ing_codigo='" & strDoc & "' "
        clsCon_Def.Ejecutar strSQL
        SaldoNC = FormatoD2(clsCon_Def.adorec_Def("saldo"))
        TotalNC = FormatoD2(clsCon_Def.adorec_Def("total"))
        FechaNC = clsCon_Def.adorec_Def("ing_fecha")
        NumNC = clsCon_Def.adorec_Def("nc")
        AsientoNC = clsCon_Def.adorec_Def("ing_numasiento")
        If SaldoNC > 0 Then
            If Fact = "" Then
                Filtro = " AND cuenta_p_c.cue_p_c_fechaemision BETWEEN '" & FechaDesde & "' AND '" & FechaHasta & "'"
            Else
                Filtro = " AND cuenta_p_c.cue_p_c_egr_codigo ='" & Fact & "' "
            End If
            strSQL = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta, cue_p_c_valor,cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d,tip_doc_cue_descripcion " & _
                     " FROM cuenta_p_c INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
                     " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo" & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " WHERE per_codigo IN ('" & strPersona & "') AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo = 'C' AND cue_p_c_pagado='0' " & _
                     Filtro & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo, cue_p_c_egr_codigo, cue_p_c_descripcion, cue_p_c_fechaemision, cue_p_c_fechapropuesta,cue_p_c_valor,tip_doc_cue_descripcion,com_ret_total HAVING round(cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0),2)!=0 " & _
                     " ORDER BY cuenta_p_c.cue_p_c_fechaemision DESC,cue_p_c_egr_codigo DESC,cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo "
            clsCon_Def.Ejecutar strSQL
            While SaldoNC > 0 And Not clsCon_Def.adorec_Def.EOF
                
                Cta = FormatoD0(clsCon_Def.adorec_Def("cue_p_c_codigo"))
                SaldoCta = FormatoD2(clsCon_Def.adorec_Def("d"))
                Desc = UCase(clsCon_Def.adorec_Def("tip_doc_cue_descripcion")) & ": " & clsCon_Def.adorec_Def("cue_p_c_egr_codigo")
                'Calcula el máximo codigo de pago para la cuenta
                strSQL = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                         " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                         "                                 AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                         "                                 AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                         " WHERE cuenta_p_c.cue_p_c_codigo= '" & Cta & "' AND pago.emp_codigo = '" & strEmpresa & "' " & _
                         " AND pago.cue_p_c_tipo = 'C'" & _
                         " GROUP BY pago.emp_codigo"
                clsCod.Ejecutar strSQL
                If clsCod.adorec_Def.EOF Then
                    maxpag = 1
                Else
                    maxpag = clsCod.adorec_Def("pag") + 1
                End If
                
                If SaldoNC >= SaldoCta Then
                    ValorPago = FormatoD2(SaldoCta)
                    Desc = Desc & " CANCELA"
                Else
                    ValorPago = FormatoD2(SaldoNC)
                    Desc = Desc & " ABONA"
                End If
                
                MsgBox "Nota de Crédito se Aplicará: " & Desc, vbInformation, "Aplicacion"
                
                strSQL = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, " & _
                         " pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "', '" & Val(Cta) & "', 'C', '" & Val(maxpag) & "', '" & FechaNC & "', '" & ValorPago & "', " & _
                         " '" & NumNC & "', '" & UCase(ClienteNC & " - " & Desc & " - NOTA DE CRÉDITO " & NumNC & " (" & strDoc & ")") & "', " & _
                         " '','" & AsientoNC & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsPag.Ejecutar strSQL, "M"
                
                
                strSQL = " SELECT cuenta_p_c.cue_p_c_codigo, cue_p_c_valor-COALESCE(com_ret_total,0)-COALESCE(sum(pag_monto),0) as d " & _
                         " FROM cuenta_p_c INNER JOIN tipo_doc_cuenta ON cuenta_p_c.tip_doc_cue_codigo=tipo_doc_cuenta.tip_doc_cue_codigo " & _
                         " LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo" & _
                         " LEFT JOIN comprobante_retencion ON cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo AND cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                         " WHERE per_codigo IN ('" & strPersona & "') AND cuenta_p_c.emp_codigo = '" & strEmpresa & "' " & _
                         " AND cuenta_p_c.cue_p_c_tipo = 'C' " & _
                         " AND cuenta_p_c.cue_p_c_codigo='" & Val(Cta) & "'" & _
                         " GROUP BY cuenta_p_c.cue_p_c_codigo, cue_p_c_valor,com_ret_total "
                clsPag.Ejecutar strSQL, "M"
                
                If FormatoD2(clsPag.adorec_Def("d")) <= 0 Then
                    strSQL = " SELECT MAX(pag_fecha) as fec " & _
                             " FROM pago " & _
                             " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & Val(Cta) & "' " & _
                             " AND emp_codigo = '" & strEmpresa & "' " & _
                             " AND pag_observacion NOT LIKE 'ANULADO%' " & _
                             " GROUP BY emp_codigo"
                    clsPag.Ejecutar strSQL, "M"
                    strSQL = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_fechapago='" & clsPag.adorec_Def("fec") & "', " & _
                             " cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, " & _
                             " cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE cue_p_c_tipo= 'C' AND cue_p_c_codigo= '" & Val(Cta) & "' " & _
                             " AND emp_codigo = '" & strEmpresa & "' "
                    clsPag.Ejecutar strSQL, "M"
                
                End If
                SaldoNC = FormatoD2(SaldoNC - ValorPago)
                clsCon_Def.adorec_Def.MoveNext
            Wend
            clsInventa.strTipo = "DCL"
            clsInventa.strIE = "I"
            clsInventa.strDoc = strDoc
            clsInventa.ModificaIng , , , , , , , , , , , , , , , FormatoD2(TotalNC - SaldoNC)
                
        End If
    Else
        MsgBox "La Nota de Crédito no se Aplicará", vbInformation, "Aplicación"
    End If
End Sub

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
    
    Set clsCostear_Def = New clsCostear
    clsCostear_Def.Inicializar adocon_ParL, adocon_ParM
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub



Public Sub NuevaCta(TipoCta As String, TipoDoc As String, Sustento As String, FechaEmi As String, FechaVen As String, Persona As String, Observ As String, Optional strSerie As String = "", Optional strNumero As String = "", Optional strAutorizacion As String = "", Optional strCaduca As String = "00/0000", Optional TotalProd As Double = 0, Optional TotalServ As Double = 0, Optional TotalProdIVA As Double = 0, Optional TotalServIVA As Double = 0, Optional codIVA As Double = 0, Optional IVA As Double = 0, Optional SubTotal0 As Double = 0, Optional baseICE As Double = 0, Optional codICE As Double = 0, Optional ICE As Double = 0, Optional Total As Double = 0, Optional Asiento As String = "")
    Dim strSQL As String
    Dim Doc As String
    Dim NumIng As Double
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    strTipoCta = TipoCta
    strTipoDoc = TipoDoc
    strSustento = Sustento
    strFechaEmision = FechaEmi
    strFechaVence = FechaVen
    strPersona = Persona
    dblTotalProd = TotalProd
    dblTotalServ = TotalServ
    dblTotalProdIVA = TotalProdIVA
    dblTotalServIVA = TotalServIVA
    dblCero = SubTotal0
    dblcodIVA = codIVA
    dblIVA = IVA
    dblBaseICE = baseICE
    dblcodICE = codICE
    dblICE = ICE
    dblTotal = Total
    Observ = UCase(Observ)
    
    strSQL = " BEGIN TRAN "
    clsCon_Aux.Ejecutar strSQL, "M"
    
    strSQL = " SELECT COALESCE(MAX(cue_p_c_codigo)+1,1) AS num " & _
             " FROM cuenta_p_c WITH (TABLOCKX)" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cue_p_c_tipo='" & strTipoCta & "' " & _
             " GROUP BY emp_codigo"
    clsCon_Aux.Ejecutar strSQL, "M"
    
    strSQL = " INSERT INTO cuenta_p_c (asi_numasiento,emp_codigo,tip_doc_cue_codigo,cod_sus_com_codigo,cue_p_c_serie,cue_p_c_numero,cue_p_c_autorizacion,cue_p_c_caduca,cue_p_c_egr_codigo,cue_p_c_codigo,cue_p_c_tipo,per_codigo, " & _
             " cue_p_c_fechaemision,cue_p_c_fechapropuesta,cue_p_c_st_cero,cue_p_c_st_prod,cue_p_c_st_serv,cod_iva_codigo,cue_p_c_iva,cue_p_c_baseice,cod_ice_codigo,cue_p_c_ice,cue_p_c_valor,cue_p_c_descripcion,cue_p_c_fechamod,cue_p_c_usumod,cue_p_c_tot_cuenta,cue_p_c_fra_cuenta) " & _
             " VALUES ('" & Asiento & "','" & strEmpresa & "','" & strTipoDoc & "','" & strSustento & "','" & strSerie & "','" & strNumero & "','" & strAutorizacion & "','" & strCaduca & "','" & FormatoD0(strSerie & strNumero) & "','" & clsCon_Aux.adorec_Def("num") & "','" & strTipoCta & "','" & strPersona & "', " & _
             " '" & strFechaEmision & "','" & strFechaVence & "'," & FormatoD2(dblCero) & "," & FormatoD2(dblTotalProdIVA) & "," & FormatoD2(dblTotalServIVA) & "," & FormatoD2(dblcodIVA) & "," & FormatoD2(dblIVA) & "," & FormatoD2(dblBaseICE) & "," & FormatoD2(dblcodICE) & "," & FormatoD2(dblICE) & "," & FormatoD2(dblTotal) & ",'" & UCase(Observ) & "', " & _
             " CURRENT_TIMESTAMP, '" & strUsuario & "',1,1) "
    clsCon_Def.Ejecutar strSQL, "M"
    
    strSQL = " COMMIT TRAN "
    clsCon_Def.Ejecutar strSQL, "M"
    
    strNoCta = clsCon_Aux.adorec_Def("num")
End Sub

Public Sub ModificarCta(Optional TipoDoc As String = "", Optional Sustento As String = "", Optional FechaEmi As String = "", Optional FechaVen As String = "", Optional Persona As String = "", Optional Observ As String = "", Optional strSerie As String = "", Optional strNumero As String = "", Optional strAutorizacion As String = "", Optional strCaduca As String = "00/0000", Optional CambiarValores As Boolean = False, Optional TotalProd As Double = 0, Optional TotalServ As Double = 0, Optional TotalProdIVA As Double = 0, Optional TotalServIVA As Double = 0, Optional codIVA As Double = 0, Optional IVA As Double = 0, Optional SubTotal0 As Double = 0, Optional baseICE As Double = 0, Optional codICE As Double = 0, Optional ICE As Double = 0, Optional Total As Double = 0, Optional Asiento As String = "")
    Dim strSQL As String
    strSQL = " UPDATE cuenta_p_c SET "
    
    If TipoDoc <> "" Then
        strSQL = strSQL & " tip_doc_cue_codigo='" & TipoDoc & "',"
        strTipoDoc = TipoDoc
    End If
    If Sustento <> "" Then
        strSQL = strSQL & " cod_sus_com_codigo='" & Sustento & "',"
        strSustento = Sustento
    End If
    If FechaEmi <> "" Then
        strSQL = strSQL & " cue_p_c_fechaemision='" & FechaEmi & "',"
        strFechaEmision = FechaEmi
    End If
    If FechaVen <> "" Then
        strSQL = strSQL & " cue_p_c_fechapropuesta='" & FechaVen & "',"
        strFechaVence = FechaVen
    End If
    If Persona <> "" Then
        strSQL = strSQL & " per_codigo='" & Persona & "',"
        strPersona = Persona
    End If
    If Observ <> "" Then
        strSQL = strSQL & " cue_p_c_descripcion='" & Observ & "',"
    End If
    If strSerie <> "" Then
        strSQL = strSQL & " cue_p_c_serie='" & strSerie & "',"
    End If
    If strNumero <> "" Then
        strSQL = strSQL & " cue_p_c_numero='" & strNumero & "',"
    End If
    If strAutorizacion <> "" Then
        strSQL = strSQL & " cue_p_c_autorizacion='" & strAutorizacion & "',"
    End If
    If strCaduca <> "00/0000" Then
        strSQL = strSQL & " cue_p_c_caduca='" & strCaduca & "',"
    End If
    If Asiento <> "" Then
        strSQL = strSQL & " asi_numasiento='" & Asiento & "',"
    End If
    
    If CambiarValores = True Then
        strSQL = strSQL & " cue_p_c_st_prod='" & TotalProdIVA & "',"
        strSQL = strSQL & " cue_p_c_st_serv='" & TotalServIVA & "',"
        strSQL = strSQL & " cue_p_c_st_cero='" & SubTotal0 & "',"
        strSQL = strSQL & " cod_iva_codigo='" & codIVA & "',"
        strSQL = strSQL & " cue_p_c_iva='" & IVA & "',"
        strSQL = strSQL & " cue_p_c_baseice='" & baseICE & "',"
        strSQL = strSQL & " cod_ice_codigo='" & codICE & "',"
        strSQL = strSQL & " cue_p_c_ice='" & ICE & "',"
        strSQL = strSQL & " cue_p_c_valor='" & Total & "',"
        
        dblTotalProd = TotalProd
        dblTotalServ = TotalServ
        dblTotalProdIVA = TotalProdIVA
        dblTotalServIVA = TotalServIVA
        dblCero = SubTotal0
        dblcodIVA = codIVA
        dblIVA = IVA
        dblBaseICE = baseICE
        dblcodICE = codICE
        dblICE = ICE
        dblTotal = Total
    End If
    strSQL = Left(strSQL, Len(strSQL) - 1)
    strSQL = strSQL & " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cue_p_c_tipo='" & strTipoCta & "' " & _
             " AND cue_p_c_codigo='" & strNoCta & "'"
    clsCon_Def.Ejecutar strSQL, "M"
        
End Sub

Public Sub IngRetencionPersonaIng(clsIngreso As clsInventario, strSerieRet As String, strNumeroRet As String, strAutorizacionRet As String, Optional strFechaRet As String = "")
    Dim strSQL As String
    Dim dblIV As Double
    Dim dblIVs
    If dblTotalProdIVA + dblTotalServIVA = 0 Then
        dblIV = 0
        dblIVs = 0
    Else
        dblIV = dblTotalProdIVA * dblIVA / (dblTotalProdIVA + dblTotalServIVA)
        dblIVs = dblTotalServIVA * dblIVA / (dblTotalProdIVA + dblTotalServIVA)
    End If
    strSQL = " INSERT INTO comprobante_retencion " & _
             " SELECT det_ingreso_ret.emp_codigo,'" & strNoCta & "','" & strTipoCta & "','" & strSerieRet & "','" & strNumeroRet & "','" & strAutorizacionRet & "', " & _
             " IIF('" & strFechaRet & "'!='','" & Format(strFechaRet, "yyyy-mm-dd") & "',ing_fecha),SUM(det_ing_ret_valor),CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
             " FROM det_ingreso_ret INNER JOIN ingreso ON det_ingreso_ret.emp_codigo=ingreso.emp_codigo AND det_ingreso_ret.tip_ing_codigo=ingreso.tip_ing_codigo AND det_ingreso_ret.ing_codigo=ingreso.ing_codigo" & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.tip_ing_codigo='" & clsIngreso.strTipo & "' " & _
             " AND ingreso.ing_codigo='" & clsIngreso.strDoc & "' GROUP BY det_ingreso_ret.emp_codigo"
    clsCon_Def.Ejecutar strSQL, "M"
    strSQL = " COMMIT TRAN "
    clsCon_Def.Ejecutar strSQL, "M"
    strSQL = " INSERT INTO det_comp_ret " & _
             " SELECT retencion.emp_codigo,'" & strNoCta & "','" & strTipoCta & "',retencion.ret_codigo,IIF(retencion.ret_gravara='SUBTOTAL','" & dblTotalProd + dblTotalServ & "',IIF(retencion.ret_gravara='SUBTOTALPRODUCTOS','" & dblTotalProd & "',IIF(retencion.ret_gravara='SUBTOTALSERVICIOS','" & dblTotalServ & "',IIF(retencion.ret_gravara='IVA','" & dblIVA & "',IIF(retencion.ret_gravara='IVAPRODUCTOS','" & dblIV & "',IIF(retencion.ret_gravara='IVASERVICIOS','" & dblIVs & "',IIF(retencion.ret_gravara='IVA0%','" & dblCero & "',IIF(retencion.ret_gravara='Total','" & dblTotal & "','0')))))))),retencion.ret_porcentaje,CURRENT_TIMESTAMP,'" & strUsuario & "' " & _
             " FROM retencion INNER JOIN persona_ret ON (retencion.ret_codigo = persona_ret.ret_codigo) " & _
             " AND (retencion.emp_codigo = persona_ret.emp_codigo) " & _
             " WHERE '" & strFechaEmision & "' BETWEEN ret_fechaini AND ret_fechafin AND persona_ret.per_codigo='" & strPersona & "' AND persona_ret.emp_codigo='" & strEmpresa & "' "
    clsCon_Def.Ejecutar strSQL, "M"
    
End Sub

Public Sub IngAsientoIng(clsAsiento As clsContable, clsIngreso As clsInventario)
    Dim strSQL As String
    Dim strCentroCosto As String
    strCentroCosto = ""
    If clsIngreso.strTipo = "COM" Then
        'cuenta contable CXP
        strSQL = " SELECT cat_p_ctaconta " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " WHERE persona.per_codigo= '" & clsIngreso.strPersona & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsCon_Def.Ejecutar strSQL
        clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("cat_p_ctaconta"), "", 0, FormatoD2(clsIngreso.dblTotal) - FormatoD2(clsIngreso.dblRetencion)
        'cuenta contable IVA COMPRAS
        If FormatoD2(clsIngreso.dblIVA) <> 0 Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='IVAC' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), "", FormatoD2(clsIngreso.dblIVA), 0
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If clsIngreso.dblTotalProd <> 0 And clsIngreso.dblTotalServ = 0 Then
            clsIngreso.dblTotalProd = clsIngreso.dblSubTotal
        ElseIf clsIngreso.dblTotalServ <> 0 And clsIngreso.dblTotalProd = 0 Then
            clsIngreso.dblTotalServ = clsIngreso.dblSubTotal
        Else
            clsIngreso.dblTotalServ = clsIngreso.dblSubTotal - clsIngreso.dblTotalProd
        End If
        If FormatoD2(clsIngreso.dblTotalProd) <> 0 Then
            strSQL = " SELECT tip_ing_ctaconta " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & clsIngreso.strTipo & "' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_ing_ctaconta"), "", FormatoD2(clsIngreso.dblTotalProd), 0
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(clsIngreso.dblTotalServ) <> 0 Then
            strSQL = " SELECT tip_ing_ctaconta2 " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & clsIngreso.strTipo & "' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_ing_ctaconta2"), "", FormatoD2(clsIngreso.dblTotalServ), 0
        End If
        'cuentas contables de RECARGOS
        strSQL = " SELECT oca_ctaconta, det_ing_c_cantidad*det_ing_c_precio as Tot " & _
                 " FROM det_ingreso_c INNER JOIN ocargos ON det_ingreso_c.emp_codigo=ocargos.emp_codigo AND det_ingreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_ingreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_ingreso_c.tip_ing_codigo='" & clsIngreso.strTipo & "' " & _
                 " AND det_ingreso_c.ing_codigo='" & clsIngreso.strDoc & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        While Not clsCon_Def.adorec_Def.EOF
            If FormatoD2(clsCon_Def.adorec_Def("Tot")) <> 0 Then
                clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("oca_ctaconta"), "", FormatoD2(clsCon_Def.adorec_Def("Tot")), 0
            End If
            clsCon_Def.adorec_Def.MoveNext
        Wend
        'cuentas contables de RETENCIONES
        strSQL = " SELECT ret_ctaconta,det_ing_ret_valor " & _
                 " FROM det_ingreso_ret INNER JOIN retencion ON retencion.emp_codigo=det_ingreso_ret.emp_codigo AND retencion.ret_codigo=det_ingreso_ret.ret_codigo " & _
                 " WHERE det_ingreso_ret.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_ingreso_ret.tip_ing_codigo='" & clsIngreso.strTipo & "' " & _
                 " AND '" & strFechaEmision & "' BETWEEN ret_fechaini AND ret_fechafin " & _
                 " AND det_ingreso_ret.ing_codigo='" & clsIngreso.strDoc & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        While Not clsCon_Def.adorec_Def.EOF
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("ret_ctaconta"), "", 0, clsCon_Def.adorec_Def("det_ing_ret_valor")
            clsCon_Def.adorec_Def.MoveNext
        Wend
    ElseIf clsIngreso.strTipo = "DCL" Then
        'cuenta contable CXC
        strSQL = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE persona.per_codigo= '" & clsIngreso.strPersona & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsCon_Def.Ejecutar (strSQL)
        clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("cat_p_ctaconta"), "", 0, FormatoD2(clsIngreso.dblTotal)
        strCentroCosto = clsCon_Def.adorec_Def("cen_cos_codigo")
        'cuenta contable IVA VENTAS
        If FormatoD2(clsIngreso.dblIVA) <> 0 And clsIngreso.booSinIva = False Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='IVAV' "
            clsCon_Def.Ejecutar (strSQL)
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), "", FormatoD2(clsIngreso.dblIVA), 0
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(clsIngreso.dblTotalProd) <> 0 Then
            strSQL = " SELECT tip_ing_ctaconta " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & clsIngreso.strTipo & "' "
            clsCon_Def.Ejecutar (strSQL)
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_ing_ctaconta"), strCentroCosto, FormatoD2(clsIngreso.dblTotalProd), 0
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(clsIngreso.dblTotalServ) <> 0 Then
            strSQL = " SELECT tip_ing_ctaconta2 " & _
                     " FROM tipo_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & clsIngreso.strTipo & "' "
            clsCon_Def.Ejecutar (strSQL)
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_ing_ctaconta2"), strCentroCosto, FormatoD2(clsIngreso.dblTotalServ), 0
        End If
        'cuentas contables de recargos
        strSQL = " SELECT oca_ctaconta, det_ing_c_cantidad*det_ing_c_precio as Tot " & _
                 " FROM det_ingreso_c INNER JOIN ocargos ON det_ingreso_c.emp_codigo=ocargos.emp_codigo AND det_ingreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_ingreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_ingreso_c.tip_ing_codigo='" & clsIngreso.strTipo & "' " & _
                 " AND det_ingreso_c.ing_codigo='" & clsIngreso.strDoc & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        While Not clsCon_Def.adorec_Def.EOF
            If FormatoD2(clsCon_Def.adorec_Def("Tot")) <> 0 Then
                clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("oca_ctaconta"), strCentroCosto, FormatoD2(clsCon_Def.adorec_Def("Tot")), 0
            End If
            clsCon_Def.adorec_Def.MoveNext
        Wend
        'cuenta contable DESCUENTO EN VENTAS
        If FormatoD2(clsIngreso.dblTotalServ) = 0 And FormatoD2(clsIngreso.dblTotalProd) = 0 And FormatoD2(clsIngreso.dblDcto) <> 0 Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='DCV' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), strCentroCosto, FormatoD2(clsIngreso.dblDcto) * (-1), 0
        End If
    End If
End Sub

Public Sub VerRet()
    frmReporte.strNumero = strNoCta
    frmReporte.strTipo = strTipoCta
    frmReporte.strReporte = "rptRetencion"
    frmReporte.Show
End Sub

Public Sub IngAsientoEgr(clsAsiento As clsContable, clsEgreso As clsInventario)
    Dim strSQL As String
    Dim secPub As String
    Dim strCentroCosto As String
    strCentroCosto = ""
    If clsEgreso.strTipo = "FAC" Then
        'cuenta contable CXC
        strSQL = " SELECT cat_p_ctaconta,cen_cos_codigo " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " INNER JOIN tipo_pedido ON persona.emp_codigo=tipo_pedido.emp_codigo " & _
                " AND persona.tip_ped_codigo=tipo_pedido.tip_ped_codigo " & _
                " WHERE persona.per_codigo= '" & clsEgreso.strPersona & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsCon_Def.Ejecutar (strSQL)
        clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("cat_p_ctaconta"), "", FormatoD2(clsEgreso.dblTotal), 0
        strCentroCosto = clsCon_Def.adorec_Def("cen_cos_codigo")
        'cuenta contable IVA VENTAS
        If FormatoD2(clsEgreso.dblIVA) <> 0 And clsEgreso.booSinIva = False Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='IVAV' "
            clsCon_Def.Ejecutar (strSQL)
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), "", 0, FormatoD2(clsEgreso.dblIVA)
        End If
        If clsEgreso.booSecPublico = False Then
            secPub = ""
        Else
            secPub = "_sp"
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(clsEgreso.dblTotalProd) <> 0 Then
            strSQL = " SELECT suc.suc_ctaconta_ventas" & secPub & " as tip_egr_ctaconta" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & strEmpresa & "' " & _
                     " AND te.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
                     " AND suc.suc_codigo= '" & strSucursal & "' "
            clsCon_Def.Ejecutar (strSQL)
            If clsCon_Def.adorec_Def.RecordCount > 0 Then
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_egr_ctaconta"), strCentroCosto, 0, FormatoD2(clsEgreso.dblTotalProd)
            Else
                strSQL = " SELECT suc.suc_ctaconta_ventas" & secPub & " as tip_egr_ctaconta" & _
                         " FROM tipo_egreso te" & _
                         " INNER JOIN sucursal suc " & _
                         " ON suc.emp_codigo=te.emp_codigo " & _
                         " WHERE te.emp_codigo='" & strEmpresa & "' " & _
                         " AND te.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
                         " AND suc.suc_codigo= '" & strSucursal2 & "' "
                clsCon_Def.Ejecutar (strSQL)
                clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_egr_ctaconta"), strCentroCosto, 0, FormatoD2(clsEgreso.dblTotalProd)
            End If
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(clsEgreso.dblTotalServ) <> 0 Then
            strSQL = " SELECT suc.suc_ctaconta_servicios" & secPub & " as tip_egr_ctaconta2" & _
                     " FROM tipo_egreso te" & _
                     " INNER JOIN sucursal suc " & _
                     " ON suc.emp_codigo=te.emp_codigo " & _
                     " WHERE te.emp_codigo='" & strEmpresa & "' " & _
                     " AND te.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
                     " AND suc.suc_codigo= '" & strSucursal & "' "
            clsCon_Def.Ejecutar (strSQL)
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_egr_ctaconta2"), strCentroCosto, 0, FormatoD2(clsEgreso.dblTotalServ)
        End If
        'cuentas contables de recargos
        strSQL = " SELECT oca_ctaconta, det_egr_c_cantidad*det_egr_c_precio as Tot " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
                 " AND det_egreso_c.egr_codigo='" & clsEgreso.strDoc & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        While Not clsCon_Def.adorec_Def.EOF
            If FormatoD2(clsCon_Def.adorec_Def("Tot")) <> 0 Then
                clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("oca_ctaconta"), strCentroCosto, 0, FormatoD2(clsCon_Def.adorec_Def("Tot"))
            End If
            clsCon_Def.adorec_Def.MoveNext
        Wend
    ElseIf clsEgreso.strTipo = "DPV" Then
        'cuenta contable CXP
        strSQL = " SELECT cat_p_ctaconta " & _
                " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo " & _
                " AND persona.cat_p_tipo=categoria_p.cat_p_tipo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                " INNER JOIN ctaconta ON categoria_p.cat_p_ctaconta= ctaconta.cta_codigo " & _
                " AND categoria_p.emp_codigo = ctaconta.emp_codigo " & _
                " WHERE persona.per_codigo= '" & clsEgreso.strPersona & "' AND persona.emp_codigo = '" & strEmpresa & "' "
        clsCon_Def.Ejecutar strSQL
        clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("cat_p_ctaconta"), "", FormatoD2(clsEgreso.dblTotal) - FormatoD2(clsEgreso.dblRetencion), 0
        'cuenta contable IVA COMPRAS
        If FormatoD2(clsEgreso.dblIVA) <> 0 Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='IVAC' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), "", 0, FormatoD2(clsEgreso.dblIVA)
        End If
        'cuenta contable FACTURAS PRODUCTOS
        If FormatoD2(clsEgreso.dblTotalProd) <> 0 Then
            strSQL = " SELECT tip_egr_ctaconta " & _
                     " FROM tipo_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & clsEgreso.strTipo & "' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_egr_ctaconta"), "", 0, FormatoD2(clsEgreso.dblTotalProd)
        End If
        'cuenta contable FACTURAS SERVICIOS
        If FormatoD2(clsEgreso.dblTotalServ) <> 0 Then
            strSQL = " SELECT tip_egr_ctaconta2 " & _
                     " FROM tipo_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & clsEgreso.strTipo & "' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("tip_egr_ctaconta2"), "", 0, FormatoD2(clsEgreso.dblTotalServ)
        End If
        'cuentas contables de RECARGOS
        strSQL = " SELECT oca_ctaconta, det_egr_c_cantidad*det_egr_c_precio as Tot " & _
                 " FROM det_egreso_c INNER JOIN ocargos ON det_egreso_c.emp_codigo=ocargos.emp_codigo AND det_egreso_c.oca_codigo=ocargos.oca_codigo " & _
                 " WHERE det_egreso_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND det_egreso_c.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
                 " AND det_egreso_c.egr_codigo='" & clsEgreso.strDoc & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        While Not clsCon_Def.adorec_Def.EOF
            If FormatoD2(clsCon_Def.adorec_Def("Tot")) <> 0 Then
                clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("oca_ctaconta"), "", 0, FormatoD2(clsCon_Def.adorec_Def("Tot"))
            End If
            clsCon_Def.adorec_Def.MoveNext
        Wend
'        'cuentas contables de RETENCIONES
'        strSql = " SELECT ret_ctaconta,det_egr_ret_valor " & _
'                 " FROM det_egreso_ret INNER JOIN retencion ON retencion.emp_codigo=det_egreso_ret.emp_codigo AND retencion.ret_codigo=det_egreso_ret.ret_codigo " & _
'                 " WHERE det_egreso_ret.emp_codigo='" & strEmpresa & "' " & _
'                 " AND det_egreso_ret.tip_egr_codigo='" & clsEgreso.strTipo & "' " & _
'                 " AND '" & strFechaEmision & "' BETWEEN ret_fechaini AND ret_fechafin " & _
'                 " AND det_egreso_ret.egr_codigo='" & clsEgreso.strDoc & "' "
'        clsCon_Def.Ejecutar strSql, "M"
'        While Not clsCon_Def.adorec_Def.EOF
'            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("ret_ctaconta"), "", clsCon_Def.adorec_Def("det_egr_ret_valor"), 0
'            clsCon_Def.adorec_Def.MoveNext
'        Wend
        'cuenta contable DESCUENTO EN COMPRAS
        If FormatoD2(clsEgreso.dblTotalServ) = 0 And FormatoD2(clsEgreso.dblTotalProd) = 0 And FormatoD2(clsEgreso.dblDcto) <> 0 Then
            strSQL = " SELECT par_texto " & _
                     " FROM parametro " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND par_codigo='DCC' "
            clsCon_Def.Ejecutar strSQL
            clsAsiento.NuevoDetAsiento clsCon_Def.adorec_Def("par_texto"), "", 0, FormatoD2(clsEgreso.dblDcto) * (-1)
        End If
    End If
End Sub

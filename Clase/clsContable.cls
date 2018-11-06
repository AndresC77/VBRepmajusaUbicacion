VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsCon_Def As New clsConsulta
Public NumAsiento As String

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
End Sub

Private Function NumeroDeNuevoAsiento(TipoAsiento As String, FechaAsiento As String) As String
    Dim maximo As String
    Dim sentencia_SQL As String
    
    'Busca el código máximo de la tabla ctaconta
    sentencia_SQL = " SELECT COALESCE(MAX(SUBSTRING(asi_numasiento,6,8)+1),1) AS numAs " & _
                    " FROM asiento WITH (TABLOCKX)" & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND asi_numasiento LIKE '" & Format(FechaAsiento, "yyyy") & TipoAsiento & "%' " & _
                    " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar sentencia_SQL, "M"
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
      maximo = clsCon_Def.adorec_Def("numAS")
      NumeroDeNuevoAsiento = Format(FechaAsiento, "yyyy") & TipoAsiento & Format(maximo, "0000000")
    Else
      maximo = 1
      NumeroDeNuevoAsiento = Format(FechaAsiento, "yyyy") & TipoAsiento & Format(maximo, "0000000")
    End If
End Function
Public Sub AnularAsientoYOtros(Motivo As String, Descripcion As String)
    Dim strSQL As String
    Dim an As Boolean
    an = False
    'Chequeo de CXC
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        strSQL = " SELECT COALESCE(SUM(pag_monto),0) as t FROM cuenta_p_c INNER JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' AND cuenta_p_c.cue_p_c_tipo='C' AND cuenta_p_c.asi_numasiento='" & NumAsiento & "'"
        clsCon_Def.Ejecutar strSQL
        If FormatoD2(clsCon_Def.adorec_Def("t")) <> 0 Then
            MsgBox "No se puede anular la cuenta ya que tienen cobros"
            Exit Sub
        End If
        If MsgBox("Asiento generado en CUENTAS POR COBRAR" & vbNewLine & "Desea Anular la Cuenta y si hubiera Cobros registrarlos como Anticipos?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularCX "C", Motivo
            an = True
        End If
    End If
    'Chequeo de CXP
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='P' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en CUENTAS POR PAGAR" & vbNewLine & "Desea Anular la Cuenta, anular la Retención y si hubiera Pagos registrarlos como Anticipos?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularCX "P", Motivo
            an = True
        End If
    End If
    'Chequeo de Doc Pago cheque postfechado
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento2='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en EFECTIVIZACION registro de Cheques PostFechados" & vbNewLine & "Desea des-efectivizar " & clsCon_Def.adorec_Def("num") & " Cobros?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            If DesEfectivizarPostfecheo(NumAsiento) = True Then
                AnularAsiento Motivo, Descripcion
                an = True
            End If
        End If
    End If
    'Chequeo de Doc Pago deposito
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en EFECTIVIZACION registro de Depósito" & vbNewLine & "Desea des-efectivizar " & clsCon_Def.adorec_Def("num") & " Cobros?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            DesEfectivizar NumAsiento
            an = True
        End If
    End If
    'Chequeo de Comprobantes de Egreso
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "'"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en COMPROBANTE DE EGRESO" & vbNewLine & "Desea Anular el Comprobante y reactivar las Cuentas por Pagar si hubiere?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularComprobanteEgreso NumAsiento, Motivo
            an = True
        End If
    End If
    'Chequeo de Notas de Credito
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM nota_d_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND tip_not_d_c='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en NOTA DE CREDITO" & vbNewLine & "Desea Anular el Comprobante y reactivar las Cuentas por Cobrar si hubiere?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularNotade "C", NumAsiento, Motivo
            an = True
        End If
    End If
    'Chequeo de Notas de Débito
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM nota_d_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND tip_not_d_c='D' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en NOTA DE DEDITO" & vbNewLine & "Desea Anular el Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularNotade "D", NumAsiento, Motivo
            an = True
        End If
    End If
    'Pagos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='P' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en PAGO" & vbNewLine & "Desea Anular el Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            AnularPagos NumAsiento, Motivo
            an = True
        End If
    End If
    'Ingresos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en INGRESO" & vbNewLine & "Desea Anular el Ingreso?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            Dim clsIngreso As New clsInventario
            clsIngreso.Inicializar AdoConn, AdoConnMaster
            clsIngreso.AnularIng , , NumAsiento, Motivo
            Set clsIngreso = Nothing
            an = True
        End If
    End If
    'Egresos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        If MsgBox("Asiento generado en EGRESO" & vbNewLine & "Desea Anular el Egreso?", vbQuestion + vbYesNo + vbDefaultButton2, "Contabilidad") = vbYes Then
            AnularAsiento Motivo, Descripcion
            Dim clsEgreso As New clsInventario
            clsEgreso.Inicializar AdoConn, AdoConnMaster
            clsEgreso.AnularEgr , , NumAsiento, Motivo
            Set clsEgreso = Nothing
            an = True
        End If
    End If
    If an = False Then
        AnularAsiento Motivo, Descripcion
    End If
End Sub
Private Sub AnularAsiento(Motivo As String, Descripcion As String)
    ModificarAsiento 0, 0, , , , "ASIENTO ANULADO" & vbNewLine & Motivo & vbNewLine & Descripcion
    AnularDetAsiento
End Sub
Private Sub AnularDetAsiento()
    Dim strSQL As String
    'Anula detalle de asiento poniendo en CERO
    strSQL = " UPDATE det_asiento " & _
             " SET det_asi_debe=0,det_asi_haber=0,det_asi_fechamod=CURRENT_TIMESTAMP,det_asi_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL, "M"
End Sub
Private Sub AnularNotade(strTipoNota As String, NumAsiento As String, Motivo As String)
    Dim strSQL As String
    Dim strCx As String
    Dim strTipo As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    strSQL = " UPDATE nota_d_c " & _
             " SET not_d_c_monto=0,not_d_c_descripcion=LEFT(CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',not_d_c_descripcion),255),not_d_c_fechamod=CURRENT_TIMESTAMP,not_d_c_usumod='" & strUsuario & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND tip_not_d_c='" & strTipoNota & "' "
    clsCon_Def.Ejecutar strSQL, "M"
    If strTipoNota = "C" Then
        strTipo = "C"
    Else
        strTipo = "P"
        strSQL = " SELECT pago.cue_p_c_codigo,pago.pag_codigo " & _
                 " FROM nota_d_c INNER JOIN pago ON nota_d_c.emp_codigo=pago.emp_codigo AND CAST(nota_d_c.not_d_c_codigo as varchar)=pago.doc_pag_codigo AND pago.cue_p_c_tipo='P' " & _
                 " INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
                 " LEFT JOIN comp_egreso ON cuenta_p_c.emp_codigo=comp_egreso.emp_codigo AND pago.doc_pag_codigo=comp_egreso.com_egr_codigo AND cuenta_p_c.per_codigo=comp_egreso.per_codigo " & _
                 " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                 " AND nota_d_c.asi_numasiento='" & NumAsiento & "' " & _
                 " AND comp_egreso.com_egr_codigo IS NULL "
        clsCon_Aux.Ejecutar strSQL
        If clsCon_Aux.adorec_Def.RecordCount > 0 Then
            While Not clsCon_Aux.adorec_Def.EOF
                strSQL = " UPDATE pago " & _
                         " SET pag_monto=0,pag_observacion=CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',pag_observacion),pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                         " AND cue_p_c_tipo='P' " & _
                         " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                strSQL = " UPDATE cuenta_p_c " & _
                         " SET cue_p_c_pagado=0,cue_p_c_fechamod=CURRENT_TIMESTAMP,cue_p_c_usumod='" & strUsuario & "' " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cue_p_c_tipo='P' " & _
                         " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                clsCon_Aux.adorec_Def.MoveNext
            Wend
        Else
            strSQL = " SELECT pago.cue_p_c_codigo,pago.pag_codigo " & _
                     " FROM nota_d_c INNER JOIN pago ON nota_d_c.emp_codigo=pago.emp_codigo AND nota_d_c.not_d_c_codigo=pago.doc_pag_codigo AND pago.cue_p_c_tipo='C' AND pago.pag_monto<0 " & _
                     " INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
                     " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND nota_d_c.asi_numasiento='" & NumAsiento & "' "
            clsCon_Aux.Ejecutar strSQL
            If clsCon_Aux.adorec_Def.RecordCount > 0 Then
                While Not clsCon_Aux.adorec_Def.EOF
                    strSQL = " UPDATE pago " & _
                             " SET pag_monto=0,pag_observacion=CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',pag_observacion),pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                             " AND cue_p_c_tipo='C' " & _
                             " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
                    clsCon_Def.Ejecutar strSQL, "M"
                    strSQL = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_pagado=0,cue_p_c_fechamod=CURRENT_TIMESTAMP,cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND cue_p_c_tipo='C' " & _
                             " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
                    clsCon_Def.Ejecutar strSQL, "M"
                    clsCon_Aux.adorec_Def.MoveNext
                Wend
            End If
        End If
    End If
    Set clsCon_Aux = Nothing
End Sub
Private Sub AnularComprobanteEgreso(NumAsiento As String, Motivo As String)
    Dim strSQL As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    'Anula detalle de asiento poniendo en CERO
    strSQL = " UPDATE comp_egreso " & _
             " SET com_egr_ch_estado='ANULADO',com_egr_descripcion=CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',com_egr_descripcion),com_egr_fechamod=CURRENT_TIMESTAMP,com_egr_usumod='" & strUsuario & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL, "M"
    strSQL = " SELECT pago.cue_p_c_codigo,pago.pag_codigo " & _
             " FROM comp_egreso INNER JOIN pago ON comp_egreso.emp_codigo=pago.emp_codigo AND comp_egreso.com_egr_codigo=pago.doc_pag_codigo AND comp_egreso.asi_numasiento=pago.asi_numasiento AND pago.cue_p_c_tipo='P' " & _
             " INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo AND comp_egreso.per_codigo=cuenta_p_c.per_codigo " & _
             " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND comp_egreso.asi_numasiento='" & NumAsiento & "' "
    clsCon_Aux.Ejecutar strSQL
    If clsCon_Aux.adorec_Def.RecordCount > 0 Then
        While Not clsCon_Aux.adorec_Def.EOF
            strSQL = " UPDATE pago " & _
                     " SET pag_monto=0,pag_observacion=CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',pag_observacion),pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                     " AND cue_p_c_tipo='P' " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            strSQL = " UPDATE cuenta_p_c " & _
                     " SET cue_p_c_pagado=0,cue_p_c_fechamod=CURRENT_TIMESTAMP,cue_p_c_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='P' " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            clsCon_Aux.adorec_Def.MoveNext
        Wend
    End If
    Set clsCon_Aux = Nothing
End Sub
Private Sub AnularPagos(NumAsiento As String, Motivo As String)
    Dim strSQL As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT pago.cue_p_c_codigo,pago.pag_codigo " & _
             " FROM pago INNER JOIN cuenta_p_c ON pago.emp_codigo=cuenta_p_c.emp_codigo AND pago.cue_p_c_codigo=cuenta_p_c.cue_p_c_codigo AND pago.cue_p_c_tipo=cuenta_p_c.cue_p_c_tipo " & _
             " WHERE pago.emp_codigo='" & strEmpresa & "' " & _
             " AND pago.asi_numasiento='" & NumAsiento & "' " & _
             " AND pago.cue_p_c_tipo='P' "
    clsCon_Aux.Ejecutar strSQL
    If clsCon_Aux.adorec_Def.RecordCount > 0 Then
        While Not clsCon_Aux.adorec_Def.EOF
            strSQL = " UPDATE pago " & _
                     " SET pag_monto=0,pag_observacion=CONCAT('ANULADO" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',pag_observacion),pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                     " AND cue_p_c_tipo='P' " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            strSQL = " UPDATE cuenta_p_c " & _
                     " SET cue_p_c_pagado=0,cue_p_c_fechamod=CURRENT_TIMESTAMP,cue_p_c_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='P' " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            clsCon_Aux.adorec_Def.MoveNext
        Wend
    End If
    Set clsCon_Aux = Nothing
End Sub
Private Function DesEfectivizarPostfecheo(NumAsiento As String) As Boolean
    Dim strSQL As String
    Dim EstadoCH As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento2='" & NumAsiento & "' " & _
             " AND asi_numasiento!='' AND asi_numasiento IS NOT NULL "
    clsCon_Aux.Ejecutar strSQL
    If clsCon_Aux.adorec_Def("num") = 0 Then
        strSQL = " SELECT doc_pag_codigo " & _
                 " FROM doc_pago " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento2='" & NumAsiento & "' "
        clsCon_Aux.Ejecutar strSQL
        While Not clsCon_Aux.adorec_Def.EOF
            strSQL = " UPDATE doc_pago " & _
                     " SET asi_numasiento2='',doc_pag_estado='GIRADO' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_pag_codigo='" & clsCon_Aux.adorec_Def("doc_pag_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            strSQL = " DELETE FROM det_doc_pago " & _
                     " WHERE emp_codigo='" & strEmpresa & "' AND det_doc_pag_n=1 " & _
                     " AND doc_pag_codigo='" & clsCon_Aux.adorec_Def("doc_pag_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            clsCon_Aux.adorec_Def.MoveNext
        Wend
        DesEfectivizarPostfecheo = True
    Else
        MsgBox "No se puede anular el asiento ya que " & clsCon_Aux.adorec_Def("num") & " documentos ya han sido depositado" & vbNewLine & "Primero anule el deposito", vbInformation, "Contabilidad"
        DesEfectivizarPostfecheo = False
    End If
    Set clsCon_Aux = Nothing
End Function
Private Sub DesEfectivizar(NumAsiento As String)
    Dim strSQL As String
    Dim EstadoCH As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    strSQL = " SELECT doc_pag_codigo,COALESCE(asi_numasiento2,'') as n " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Aux.Ejecutar strSQL
    While Not clsCon_Aux.adorec_Def.EOF
        If Trim(clsCon_Aux.adorec_Def("n")) = "" Then
            EstadoCH = "GIRADO"
        Else
            EstadoCH = "POSTFECHADO"
        End If
        strSQL = " UPDATE doc_pago " & _
                 " SET asi_numasiento='',doc_pag_estado='" & EstadoCH & "' " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND doc_pag_codigo='" & clsCon_Aux.adorec_Def("doc_pag_codigo") & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        clsCon_Aux.adorec_Def.MoveNext
    Wend
    Set clsCon_Aux = Nothing
End Sub

Public Sub AnularCX(strTipo As String, Motivo As String, Optional Egreso As String)
    Dim strSQL As String
    Dim strCx As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    'Ver CXC o XP
    If NumAsiento <> "" Then
        strSQL = " SELECT cue_p_c_codigo,cue_p_c_tipo,cue_p_c_valor,cue_p_c_egr_codigo " & _
                 " FROM cuenta_p_c " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & NumAsiento & "' " & _
                 " AND cue_p_c_tipo='" & strTipo & "' "
        clsCon_Aux.Ejecutar strSQL
        strCx = clsCon_Aux.adorec_Def("cue_p_c_codigo")
        
    Else
        strSQL = " SELECT cue_p_c_codigo,cue_p_c_tipo,cue_p_c_valor,cue_p_c_egr_codigo " & _
                 " FROM cuenta_p_c " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & NumAsiento & "' " & _
                 " AND cue_p_c_egr_codigo='" & Egreso & "'" & _
                 " AND cue_p_c_tipo='" & strTipo & "' " & _
                 " AND UCASE(cue_p_c_descripcion) LIKE 'FACTURA # " & Left(Egreso, Len(Egreso) - 7) & "%" & Format(Right(Egreso, 7) + 0, "000000") & " -%'"
        clsCon_Aux.Ejecutar strSQL
        If clsCon_Aux.adorec_Def.RecordCount > 0 Then
            strCx = clsCon_Aux.adorec_Def("cue_p_c_codigo")
        Else
            strCx = "-"
        End If
        
    End If
    Set clsCon_Aux = Nothing
    'Anula CXC o XP
    If strCx <> "-" Then
        strSQL = " UPDATE cuenta_p_c " & _
                 " SET cue_p_c_pagado=1,cue_p_c_descripcion=CONCAT('CX" & strTipo & " ANULADA" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',cue_p_c_descripcion,'" & vbNewLine & "VALOR=',cue_p_c_valor)," & _
                 " cue_p_c_st_prod=0,cue_p_c_st_serv=0,cue_p_c_st_cero=0,cue_p_c_iva=0,cue_p_c_valor=0,cue_p_c_fechamod=CURRENT_TIMESTAMP,cue_p_c_usumod='" & strUsuario & "'" & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND asi_numasiento='" & NumAsiento & "' " & _
                 " AND cue_p_c_codigo='" & strCx & "' " & _
                 " AND cue_p_c_tipo='" & strTipo & "' "
        clsCon_Def.Ejecutar strSQL, "M"
        QuitarPago Motivo, strCx, strTipo
    End If
    
End Sub
Private Sub QuitarPago(Motivo As String, strCx As String, strTipo As String)
    Dim strSQL As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    'Anula Cobros
    strSQL = " UPDATE pago " & _
             " SET pag_observacion=CONCAT('CX" & strTipo & " ANULADA" & vbNewLine & "','" & Motivo & "" & vbNewLine & "',pag_observacion,'" & vbNewLine & "VALOR=',pag_monto), " & _
             " pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cue_p_c_tipo='" & strTipo & "' " & _
             " AND cue_p_c_codigo='" & strCx & "' "
    clsCon_Def.Ejecutar strSQL, "M"
    strSQL = " SELECT DISTINCT doc_pag_codigo,pag_codigo,asi_numasiento " & _
             " FROM pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cue_p_c_tipo='" & strTipo & "' " & _
             " AND cue_p_c_codigo='" & strCx & "' "
    clsCon_Aux.Ejecutar strSQL
    If strTipo = "C" Then
        If clsCon_Aux.adorec_Def.RecordCount > 0 Then
            While Not clsCon_Aux.adorec_Def.EOF
                LiberarDocPago clsCon_Aux.adorec_Def("doc_pag_codigo"), clsCon_Aux.adorec_Def("asi_numasiento")
                clsCon_Aux.adorec_Def.MoveNext
            Wend
        End If
    Else
        If clsCon_Aux.adorec_Def.RecordCount > 0 Then
            While Not clsCon_Aux.adorec_Def.EOF
                strSQL = " UPDATE pago " & _
                         " SET doc_pag_codigo='',pag_monto=0 " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                         " AND cue_p_c_tipo='" & strTipo & "' " & _
                         " AND cue_p_c_codigo='" & strCx & "' "
                clsCon_Def.Ejecutar strSQL, "M"
                clsCon_Aux.adorec_Def.MoveNext
            Wend
        End If
            'ELIMINAR RETENCIONES
            strSQL = " UPDATE comprobante_retencion " & _
                     " SET com_ret_total=0, " & _
                     " com_ret_fechamod=CURRENT_TIMESTAMP,com_ret_usumod='" & strUsuario & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='" & strTipo & "' " & _
                     " AND cue_p_c_codigo='" & strCx & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            strSQL = " UPDATE det_comp_ret " & _
                     " SET det_com_ret_valor=0, " & _
                     " det_com_ret_fechamod=CURRENT_TIMESTAMP,det_com_ret_usumod='" & strUsuario & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='" & strTipo & "' " & _
                     " AND cue_p_c_codigo='" & strCx & "' "
            clsCon_Def.Ejecutar strSQL, "M"
    End If
    Set clsCon_Aux = Nothing
End Sub
Private Sub LiberarDocPago(DocPagCodigo As String, AsiNumero As String)
    Dim strSQL As String
    Dim clsCon_Aux As New clsConsulta
    clsCon_Aux.Inicializar AdoConn, AdoConnMaster
    'Anula Cobros
    strSQL = " SELECT cue_p_c_codigo,pag_codigo " & _
             " FROM pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cue_p_c_tipo='C' " & _
             " AND asi_numasiento='" & AsiNumero & "' " & _
             " AND doc_pag_codigo='" & DocPagCodigo & "' "
    clsCon_Aux.Ejecutar strSQL
    If clsCon_Aux.adorec_Def.RecordCount > 0 Then
        While Not clsCon_Aux.adorec_Def.EOF
            strSQL = " UPDATE pago " & _
                     " SET doc_pag_codigo='',pag_monto=0,pag_observacion=CONCAT('DESAPLICACION" & vbNewLine & "',pag_observacion), " & _
                     " pag_fechamod=CURRENT_TIMESTAMP,pag_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND pag_codigo='" & clsCon_Aux.adorec_Def("pag_codigo") & "' " & _
                     " AND cue_p_c_tipo='C' " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            strSQL = " UPDATE cuenta_p_c " & _
                     " SET cue_p_c_pagado=0 " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='C' " & _
                     " AND cue_p_c_valor<>0 " & _
                     " AND cue_p_c_codigo='" & clsCon_Aux.adorec_Def("cue_p_c_codigo") & "' "
            clsCon_Def.Ejecutar strSQL, "M"
            clsCon_Aux.adorec_Def.MoveNext
        Wend
    End If
    strSQL = " UPDATE doc_pago " & _
             " SET doc_pag_anticipo=1, " & _
             " doc_pag_fechamod=CURRENT_TIMESTAMP,doc_pag_usumod='" & strUsuario & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND doc_pag_codigo='" & DocPagCodigo & "' "
    clsCon_Def.Ejecutar strSQL, "M"
    Set clsCon_Aux = Nothing
End Sub
Public Sub NuevoAsiento(TipoAsiento As String, Fecha_Asiento As String, Revisado As Integer, Mayorizado As Integer, Total_Asiento As Double, Asiento_Descripcion As String, Optional NuevoNumero As Boolean = True)
    Dim sentencia_SQL As String
    Dim Num_Asien As String
    Dim strSQL As String
    Dim clsAux As clsConsulta
    Set clsAux = New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    Inicializar AdoConn, AdoConnMaster
    clsCon_Def.Ejecuto = False
    While clsCon_Def.Ejecuto = False
        
        strSQL = " BEGIN TRAN "
        clsAux.Ejecutar strSQL, "M"
        
        If NuevoNumero = True Then
            Num_Asien = NumeroDeNuevoAsiento(TipoAsiento, Fecha_Asiento)
        Else
            Num_Asien = NumAsiento
        End If
        
        sentencia_SQL = " INSERT INTO asiento (asi_numasiento,emp_codigo, " & _
                        " asi_fecha, asi_revisado, asi_mayorizado, asi_totaldebe, " & _
                        " asi_totalhaber, asi_descripcion, asi_fechamod, asi_usumod) " & _
                        " VALUES ('" & Num_Asien & "', " & _
                        " '" & strEmpresa & "', '" & Format(Fecha_Asiento, "yyyy-mm-dd") & "', " & _
                        " '" & Revisado & "','" & Mayorizado & "', " & _
                        " " & Format(Total_Asiento, "##0.00") & ", " & _
                        " " & Format(Total_Asiento, "##0.00") & ", " & _
                        " '" & UCase(Asiento_Descripcion) & "', " & _
                        " CURRENT_TIMESTAMP, '" & strUsuario & "')"
        'MsgBox Sentencia_Sql
        
        clsCon_Def.Ejecutar sentencia_SQL, "M"
        strSQL = " COMMIT TRAN "
        clsAux.Ejecutar strSQL, "M"
    Wend
    
    
    'asigna el valor del numero de asiento para el detalle
    NumAsiento = Num_Asien
End Sub
Public Sub ModificarAsiento(Total_debe As Double, Total_haber As Double, Optional Fecha_Asiento As String = "", Optional Revisado As Integer = -1, Optional Mayorizado As Integer = -1, Optional Asiento_Descripcion As String = "")
    Dim sentencia_SQL As String
    sentencia_SQL = " UPDATE asiento SET asi_totaldebe='" & Total_debe & "', asi_totalhaber='" & Total_haber & "' "
    If Fecha_Asiento <> "" Then
        sentencia_SQL = sentencia_SQL & ", asi_fecha='" & Fecha_Asiento & "' "
    End If
    If Revisado <> -1 Then
        sentencia_SQL = sentencia_SQL & ", asi_revisado = '" & Revisado & "' "
    End If
    If Mayorizado <> -1 Then
        sentencia_SQL = sentencia_SQL & ", asi_mayorizado = '" & Mayorizado & "' "
    End If
    If Asiento_Descripcion <> "" Then
        sentencia_SQL = sentencia_SQL & ", asi_descripcion = '" & UCase(Asiento_Descripcion) & "' "
    End If
    
    sentencia_SQL = sentencia_SQL & " WHERE asi_numasiento ='" & NumAsiento & _
                    "' AND emp_codigo='" & strEmpresa & "' "
    'MsgBox Sentencia_Sql
    
    clsCon_Def.Ejecutar sentencia_SQL, "M"
End Sub

Public Sub NuevoDetAsiento(Cuenta_Conta As String, Centro_Costo As String, DEBE As Double, HABER As Double)
    Dim sentencia_SQL As String
    
    'No graba detalles de asiento con cero en el debe y en el haber
    If DEBE = HABER And DEBE = 0 Then Exit Sub
    If Trim(Cuenta_Conta) = "" Or (DEBE = 0 And HABER = 0) Then Exit Sub
    sentencia_SQL = " SELECT COALESCE(COUNT(*),0) as num " & _
                    " FROM det_asiento " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND asi_numasiento='" & NumAsiento & "' " & _
                    " AND cta_codigo='" & Cuenta_Conta & "' " & _
                    " AND cen_cos_codigo='" & Centro_Costo & "'"
    clsCon_Def.Ejecutar sentencia_SQL
    If clsCon_Def.adorec_Def("num") = 0 Then
    sentencia_SQL = " INSERT INTO det_asiento(emp_codigo, asi_numasiento, cta_codigo,cen_cos_codigo, " & _
                    " det_asi_debe, det_asi_haber,det_asi_fechamod, det_asi_usumod) " & _
                    " VALUES ('" & strEmpresa & "','" & NumAsiento & "'," & _
                    " '" & Cuenta_Conta & "','" & Centro_Costo & "', " & _
                    " " & Format(DEBE, "##0.00") & ", " & _
                    " " & Format(HABER, "##0.00") & ", " & _
                    " CURRENT_TIMESTAMP, '" & strUsuario & "')"
    Else
        sentencia_SQL = " UPDATE det_asiento " & _
                        " SET det_asi_debe=det_asi_debe+'" & DEBE & "', " & _
                        " det_asi_haber=det_asi_haber+'" & HABER & "' " & _
                        " WHERE emp_codigo='" & strEmpresa & "' " & _
                        " AND asi_numasiento='" & NumAsiento & "' " & _
                        " AND cta_codigo='" & Cuenta_Conta & "' " & _
                        " AND cen_cos_codigo='" & Centro_Costo & "'"
    End If
    clsCon_Def.Ejecutar sentencia_SQL, "M"
End Sub

Public Sub EliminarAsiento(Cabecera As Boolean, Detalle As Boolean)
    Dim sentencia_SQL As String
    If Detalle = True Or Cabecera = True Then
        sentencia_SQL = " DELETE FROM det_asiento WHERE asi_numasiento ='" & NumAsiento & "' AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar sentencia_SQL, "M"
    End If
    
    If Cabecera = True Then
        sentencia_SQL = " DELETE FROM asiento WHERE asi_numasiento ='" & NumAsiento & "' AND emp_codigo='" & strEmpresa & "'"
        clsCon_Def.Ejecutar sentencia_SQL, "M"
    End If
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub

Public Function VerMasDatos() As String
    Dim Mensaje As String
    Dim strSQL As String
    Mensaje = ""
    'Chequeo de CXC
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Cuentas por Cobrar" & vbNewLine
    End If
    'Chequeo de CXP
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='P' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Cuentas por Pagar" & vbNewLine
    End If
    'Chequeo de Doc Pago Cheque postfechado
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento2='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado de Cheques PostFechados (" & clsCon_Def.adorec_Def("num") & " documentos)" & vbNewLine
    End If
    'Chequeo de Doc Pago Deposito a banco
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM doc_pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Efectivizacion (" & clsCon_Def.adorec_Def("num") & " documentos)" & vbNewLine
    End If
    'Chequeo de Comprobantes de Egreso
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM comp_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "'"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por Comprobante de Egreso" & vbNewLine
    End If
    'Chequeo de Notas de Credito
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM nota_d_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND tip_not_d_c='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por Nota de Crédito" & vbNewLine
    End If
    'Chequeo de Notas de Débito
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM nota_d_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND tip_not_d_c='D' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por Nota de Débito" & vbNewLine
    End If
    'Chequeo de Pagos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='P' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por un pago de una Cuenta por Pagar" & vbNewLine
    End If
    'Chequeo de Cobros
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM pago " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por una efectivizacion de un cobro de una Cuenta por Cobrar" & vbNewLine
    End If
    'Ingresos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por un Ingreso de Mercaderia" & vbNewLine
    End If
    'Egresos
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado por un Egreso de Mercaderia" & vbNewLine
    End If
    
    'RRHH
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Recursos Humanos" & vbNewLine
    End If
    'RRHH
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM descuento " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " OR asi_numasiento2='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Recursos Humanos" & vbNewLine
    End If
    'ActivoFijo
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM det_activo_fijo " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = Mensaje & "Asiento generado en Activo Fijo" & vbNewLine
    End If
    
    
    
    VerMasDatos = Mensaje
End Function
Public Function Eliminable() As Boolean
    Dim Mensaje As Boolean
    Dim strSQL As String
    Mensaje = True
    'Chequeo de CXC
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = False
    Else
        'Chequeo de CXP
        strSQL = " SELECT (count(*)) as num " & _
             " FROM cuenta_p_c " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND asi_numasiento='" & NumAsiento & "' " & _
             " AND cue_p_c_tipo='P' "
        clsCon_Def.Ejecutar strSQL
        If clsCon_Def.adorec_Def("num") > 0 Then
            Mensaje = False
        Else
            'Chequeo de Doc Pago deposito y cheques postfechados
            strSQL = " SELECT COALESCE(count(*),0) as num " & _
                     " FROM doc_pago " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND (asi_numasiento='" & NumAsiento & "' OR asi_numasiento2='" & NumAsiento & "') "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def("num") > 0 Then
                Mensaje = False
            Else
                'Chequeo de Comprobantes de Egreso
                strSQL = " SELECT COALESCE(count(*),0) as num " & _
                         " FROM comp_egreso " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND asi_numasiento='" & NumAsiento & "'"
                clsCon_Def.Ejecutar strSQL
                If clsCon_Def.adorec_Def("num") > 0 Then
                    Mensaje = False
                Else
                    'Chequeo de Notas de Credito
                    strSQL = " SELECT COALESCE(count(*),0) as num " & _
                             " FROM nota_d_c " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND asi_numasiento='" & NumAsiento & "' " & _
                             " AND tip_not_d_c='C' "
                    clsCon_Def.Ejecutar strSQL
                    If clsCon_Def.adorec_Def("num") > 0 Then
                        Mensaje = False
                    Else
                        'Chequeo de Notas de Débito
                        strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                 " FROM nota_d_c " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND asi_numasiento='" & NumAsiento & "' " & _
                                 " AND tip_not_d_c='D' "
                        clsCon_Def.Ejecutar strSQL
                        If clsCon_Def.adorec_Def("num") > 0 Then
                            Mensaje = False
                        Else
                            'Chequeo de Notas de Credito
                            strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                     " FROM pago " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                     " AND asi_numasiento='" & NumAsiento & "' "
                            clsCon_Def.Ejecutar strSQL
                            If clsCon_Def.adorec_Def("num") > 0 Then
                                Mensaje = False
                            Else
                                'Ingresos
                                strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                         " FROM ingreso " & _
                                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                                         " AND ing_numasiento='" & NumAsiento & "' "
                                clsCon_Def.Ejecutar strSQL
                                If clsCon_Def.adorec_Def("num") > 0 Then
                                    Mensaje = False
                                Else
                                    'Egresos
                                    strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                             " FROM egreso " & _
                                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                                             " AND egr_numasiento='" & NumAsiento & "' "
                                    clsCon_Def.Ejecutar strSQL
                                    If clsCon_Def.adorec_Def("num") > 0 Then
                                        Mensaje = False
                                    Else
                                        Mensaje = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    Eliminable = Mensaje
End Function

Public Function Modificable(strCta As String) As Boolean
    Dim Mensaje As Boolean
    Dim strSQL As String
    Mensaje = True
    'Chequeo de CXC
    strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c INNER JOIN det_asiento ON cuenta_p_c.emp_codigo=det_asiento.emp_codigo " & _
             " AND cuenta_p_c.asi_numasiento=det_asiento.asi_numasiento " & _
             " AND cuenta_p_c.cue_p_c_valor=det_asiento.det_asi_debe " & _
             " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
             " AND cuenta_p_c.asi_numasiento='" & NumAsiento & "' " & _
             " AND det_asiento.cta_codigo='" & strCta & "' " & _
             " AND cue_p_c_tipo='C' "
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def("num") > 0 Then
        Mensaje = False
    Else
        'Chequeo de CXP
        strSQL = " SELECT COALESCE(count(*),0) as num " & _
             " FROM cuenta_p_c INNER JOIN det_asiento ON cuenta_p_c.emp_codigo=det_asiento.emp_codigo " & _
             " AND cuenta_p_c.asi_numasiento=det_asiento.asi_numasiento " & _
             " AND cuenta_p_c.cue_p_c_valor=det_asiento.det_asi_haber " & _
             " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
             " AND cuenta_p_c.asi_numasiento='" & NumAsiento & "' " & _
             " AND det_asiento.cta_codigo='" & strCta & "' " & _
             " AND cue_p_c_tipo='P' "
        clsCon_Def.Ejecutar strSQL
        If clsCon_Def.adorec_Def("num") > 0 Then
            Mensaje = False
        Else
            'Chequeo de Doc Pago cheques postfechados
            strSQL = " SELECT COALESCE(count(*),0) as num " & _
                     " FROM doc_pago INNER JOIN det_asiento ON doc_pago.emp_codigo=det_asiento.emp_codigo " & _
                     " AND doc_pago.asi_numasiento2=det_asiento.asi_numasiento " & _
                     " AND (doc_pago.doc_pag_valor=det_asiento.det_asi_debe OR doc_pago.doc_pag_valor=det_asiento.det_asi_haber) " & _
                     " WHERE doc_pago.emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_pago.asi_numasiento2='" & NumAsiento & "' " & _
                     " AND det_asiento.cta_codigo='" & strCta & "' "
            clsCon_Def.Ejecutar strSQL
            If clsCon_Def.adorec_Def("num") > 0 Then
                Mensaje = False
            Else
                'Chequeo de Doc Pago deposito al banco
                strSQL = " SELECT COALESCE(count(*),0) as num " & _
                         " FROM doc_pago INNER JOIN det_asiento ON doc_pago.emp_codigo=det_asiento.emp_codigo " & _
                         " AND doc_pago.asi_numasiento=det_asiento.asi_numasiento " & _
                         " AND doc_pago.doc_pag_valor=det_asiento.det_asi_debe " & _
                         " WHERE doc_pago.emp_codigo='" & strEmpresa & "' " & _
                         " AND doc_pago.asi_numasiento='" & NumAsiento & "' " & _
                         " AND det_asiento.cta_codigo='" & strCta & "' "
                clsCon_Def.Ejecutar strSQL
                If clsCon_Def.adorec_Def("num") > 0 Then
                    Mensaje = False
                Else
                    'Chequeo de Comprobantes de Egreso
                    strSQL = " SELECT COALESCE(count(*),0) as num " & _
                             " FROM comp_egreso INNER JOIN det_asiento ON comp_egreso.emp_codigo=det_asiento.emp_codigo " & _
                             " AND comp_egreso.asi_numasiento=det_asiento.asi_numasiento " & _
                             " AND comp_egreso.com_egr_ch_valor=det_asiento.det_asi_haber " & _
                             " WHERE comp_egreso.emp_codigo='" & strEmpresa & "' " & _
                             " AND comp_egreso.asi_numasiento='" & NumAsiento & "' " & _
                             " AND det_asiento.cta_codigo='" & strCta & "' "
                    clsCon_Def.Ejecutar strSQL
                    If clsCon_Def.adorec_Def("num") > 0 Then
                        Mensaje = False
                    Else
                        'Chequeo de Notas de Credito
                        strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                 " FROM nota_d_c INNER JOIN det_asiento ON nota_d_c.emp_codigo=det_asiento.emp_codigo " & _
                                 " AND nota_d_c.asi_numasiento=det_asiento.asi_numasiento " & _
                                 " AND nota_d_c.not_d_c_monto=det_asiento.det_asi_haber " & _
                                 " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                                 " AND nota_d_c.asi_numasiento='" & NumAsiento & "' " & _
                                 " AND det_asiento.cta_codigo='" & strCta & "' " & _
                                 " AND tip_not_d_c='C' "
                        clsCon_Def.Ejecutar strSQL
                        If clsCon_Def.adorec_Def("num") > 0 Then
                            Mensaje = False
                        Else
                            'Chequeo de Notas de Débito
                            strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                     " FROM nota_d_c INNER JOIN det_asiento ON nota_d_c.emp_codigo=det_asiento.emp_codigo " & _
                                     " AND nota_d_c.asi_numasiento=det_asiento.asi_numasiento " & _
                                     " AND nota_d_c.not_d_c_monto=det_asiento.det_asi_haber " & _
                                     " WHERE nota_d_c.emp_codigo='" & strEmpresa & "' " & _
                                     " AND nota_d_c.asi_numasiento='" & NumAsiento & "' " & _
                                     " AND det_asiento.cta_codigo='" & strCta & "' " & _
                                     " AND tip_not_d_c='D' "
                            clsCon_Def.Ejecutar strSQL
                            If clsCon_Def.adorec_Def("num") > 0 Then
                                Mensaje = False
                            Else
                                'pagos
                                strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                         " FROM pago INNER JOIN det_asiento ON pago.emp_codigo=det_asiento.emp_codigo " & _
                                         " AND pago.asi_numasiento=det_asiento.asi_numasiento " & _
                                         " AND pago.pag_monto=det_asiento.det_asi_haber " & _
                                         " WHERE pago.emp_codigo='" & strEmpresa & "' " & _
                                         " AND pago.asi_numasiento='" & NumAsiento & "' " & _
                                         " AND det_asiento.cta_codigo='" & strCta & "' " & _
                                         " AND pago.cue_p_c_tipo='P' "
                                clsCon_Def.Ejecutar strSQL
                                If clsCon_Def.adorec_Def("num") > 0 Then
                                    Mensaje = False
                                Else
                                    'cobros
                                    strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                             " FROM pago INNER JOIN det_asiento ON pago.emp_codigo=det_asiento.emp_codigo " & _
                                             " AND pago.asi_numasiento=det_asiento.asi_numasiento " & _
                                             " AND pago.pag_monto=det_asiento.det_asi_debe " & _
                                             " WHERE pago.emp_codigo='" & strEmpresa & "' " & _
                                             " AND pago.asi_numasiento='" & NumAsiento & "' " & _
                                             " AND det_asiento.cta_codigo='" & strCta & "' " & _
                                             " AND pago.cue_p_c_tipo='C' "
                                    clsCon_Def.Ejecutar strSQL
                                    If clsCon_Def.adorec_Def("num") > 0 Then
                                        Mensaje = False
                                    Else
                                        'Ingresos
                                        strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                                 " FROM ingreso " & _
                                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                 " AND ing_numasiento='" & NumAsiento & "' "
                                        clsCon_Def.Ejecutar strSQL
                                        If clsCon_Def.adorec_Def("num") > 0 Then
                                            Mensaje = False
                                        Else
                                            'Egresos
                                            strSQL = " SELECT COALESCE(count(*),0) as num " & _
                                                     " FROM egreso " & _
                                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                                     " AND egr_numasiento='" & NumAsiento & "' "
                                            clsCon_Def.Ejecutar strSQL
                                            If clsCon_Def.adorec_Def("num") > 0 Then
                                                Mensaje = False
                                            Else
                                                Mensaje = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    Modificable = Mensaje
End Function

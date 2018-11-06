VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCostear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public clsCon_Def As clsConsulta
Public booPausa As Boolean
Private FD As String
Private fh As String

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
End Sub
Public Sub Recostear(desde As String, Optional hasta As String = "", Optional produc As String = "%", Optional siCostear As Boolean = False, Optional Saltar As Long = 0, Optional FechaHoraPausa As String = "")
    Dim strSql As String
    Dim strPrd As String
    Dim tdoc As String
    Dim Doc As String
    Dim docfecha As String
    Dim prd_codigo As String
    Dim egr As Double
    Dim ing As Double
    Dim exi As Double
    Dim d As String
    Dim prec As Double
    Dim PC As Double
    Dim PPP As Double
    Dim TotalRegistros As Long
    Dim ActRegistros As Long
    Dim booTerminar As Boolean
    booTerminar = False
    FD = desde
    fh = hasta
    If siCostear = True Then
        CreaKardex produc, desde, hasta
        booPausa = False
        'EN LOS EGRESOS POR TRANSFERENCIA SE DEBERIA RECOSTEAR EL PRODUCTO DEL INGRESOS PARA MANTENER EL COSTO
        strSql = " SELECT tdoc,doc,docfecha,prd_codigo,egr,ing,exi,d,prec,PC," & _
                 " IIF(tdoc in ('AAU','IGR','EGR','COM','IIM')," & _
                    " IIF(tdoc='AAU','A'," & _
                    " IIF(tdoc='IGR','B'," & _
                    " IIF(tdoc='EGR','C'," & _
                    " IIF(tdoc='COM','D','E'))))," & _
                 " IIF(tdoc IN ('BAU','ETN','ITN','GRE','DRE')," & _
                    " IIF(tdoc='BAU','F'," & _
                    " IIF(tdoc='ETN','G'," & _
                    " IIF(tdoc='ITN','H'," & _
                    " IIF(tdoc='GRE','I','J'))))," & _
                 " IIF(tdoc IN('ERE','IRE','NRP','DPR','ETR')," & _
                    " IIF(tdoc='ERE','K'," & _
                    " IIF(tdoc='IRE','L'," & _
                    " IIF(tdoc='NRP','M'," & _
                    " IIF(tdoc='DPR','N','O'))))," & _
                 " IIF(tdoc IN('ITR','FAC','NOT','DCL','DPV','ICA','ECA')," & _
                    " IIF(tdoc='ITR','P'," & _
                    " IIF(tdoc IN ('FAC','NOT'),'Q'," & _
                    " IIF(tdoc='DCL','R'," & _
                    " IIF(tdoc='DPV','S'," & _
                    " IIF(tdoc='ICA','T'," & _
                    " IIF(tdoc='ECA','U','Z')))))),'Z')))) as dd "
        strSql = strSql & " FROM #kardex" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " WHERE egr!=0 or ing!=0 or exi!=0 or (tdoc='AAU' AND doc=10010000001) or LEFT(tdoc,5)='SALDO'" & _
                 " ORDER BY IIF(prd_codigo='7704170724171',1,2),prd_codigo,docfecha,dd,d,tdoc,doc "
        clsCon_Def.Ejecutar strSql
        strPrd = ""
        PPP = 0
        TotalRegistros = clsCon_Def.adorec_Def.RecordCount
        ActRegistros = 0
        If Saltar <> 0 Then
            clsCon_Def.adorec_Def.AbsolutePosition = Saltar
            ActRegistros = Saltar
        End If
        While booTerminar = False And Not clsCon_Def.adorec_Def.EOF
                DoEvents
                If FechaHoraPausa <> "" Then
                    If Now > FechaHoraPausa Then
                        booPausa = True
                    End If
                End If
                If booPausa = True Then
                    FechaHoraPausa = InputBox("Fecha hora nueva pausa:", "Costeo", FechaHoraPausa)
                    
                    If MsgBox("COSTEO PAUSADO" & vbNewLine & _
                    clsCon_Def.adorec_Def.AbsolutePosition & "/" & clsCon_Def.adorec_Def.RecordCount & vbNewLine & _
                    "DESEA CONTINUAR?", vbYesNo + vbQuestion, "Pausa") = vbNo Then
                        booTerminar = True
                    End If
                    booPausa = False
                End If
                If clsCon_Def.adorec_Def("tdoc") = "SALDO INICIAL A LA FECHA" And Saltar <> 0 Then
                    Saltar = 0
                End If
                If Saltar = 0 Then
                    tdoc = Trim(clsCon_Def.adorec_Def("tdoc"))
                    Doc = clsCon_Def.adorec_Def("doc")
                    docfecha = clsCon_Def.adorec_Def("docfecha")
                    prd_codigo = clsCon_Def.adorec_Def("prd_codigo")
'                    If prd_codigo <> "7704170724171" Then
'                        'MsgBox "7704170724171"
'                    End If
                    egr = clsCon_Def.adorec_Def("egr")
                    ing = clsCon_Def.adorec_Def("ing")
                    d = clsCon_Def.adorec_Def("d")
                    prec = clsCon_Def.adorec_Def("prec")
                    If tdoc = "SALDO INICIAL A LA FECHA" Then
                        PC = UltimoPPP(prd_codigo, docfecha)
                    Else
                        PC = clsCon_Def.adorec_Def("PC")
                    End If
                    If strPrd <> prd_codigo Then
                        PPP = 0
                        exi = clsCon_Def.adorec_Def("exi")
                        strPrd = prd_codigo
                        PPP = UltimoPPP(prd_codigo, desde)
                    End If
                    PPP = CalculaPPP(tdoc, Doc, docfecha, prd_codigo, egr, ing, exi, d, prec, PC, PPP)
                    exi = exi + ing - egr
                End If
            clsCon_Def.adorec_Def.MoveNext
            If Saltar = 0 Then
                If Not clsCon_Def.adorec_Def.EOF Then
                    If strPrd <> clsCon_Def.adorec_Def("prd_codigo") Then
                        ActualizaCosto prd_codigo, PPP
                    End If
                Else
                    ActualizaCosto prd_codigo, PPP
                End If
            End If
            mdiPrincipal.StatusBar.Panels(3).Text = ActRegistros & "/" & TotalRegistros
            mdiPrincipal.StatusBar.Refresh
            ActRegistros = ActRegistros + 1
        Wend
'        MsgBox ActRegistros & "/" & TotalRegistros
        strSql = " EXEC Sp_Drop_Table_if_Exist '##ProdCosteado" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "'"
        clsCon_Def.Ejecutar strSql
        
        strSql = " EXEC Sp_Drop_Table_if_Exist '#kardex" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "'"
        clsCon_Def.Ejecutar strSql
    End If
    'mdiPrincipal.StatusBar.Panels(3).Text = ""
End Sub
Private Function CalculaPPP(tdoc As String, Doc As String, docfecha As String, prd_codigo As String, egr As Double, ing As Double, exi As Double, d As String, prec As Double, PC As Double, PPP As Double) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    Dim aux As Double
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = "SELECT * FROM parametro"
    If d = 1 Then
        If tdoc = "COM" Or tdoc = "IIM" Or (tdoc = "AAU" And Doc = "10010000001") Then 'Compras, Importaciones, produccion
            If (exi + ing) <> 0 Then
                PPP = FormatoD4((PPP * exi + ing * prec) / (exi + ing))
            Else
                PPP = prec
            End If
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "ITN" Then 'INGRESO POR TRANSFORMACION PRODUCTO TERMINADO
            prec = FormatoD4(CostoTransformacion(Doc))
            If (exi + ing) <> 0 Then
                PPP = FormatoD4((PPP * exi + ing * prec) / (exi + ing))
            Else
                PPP = prec
            End If
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "', " & _
                     " det_ing_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE ingreso " & _
                     " SET ing_subtotal='" & prec * ing & "' " & _
                     ",ing_total='" & prec * ing & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' "
        ElseIf tdoc = "AAU" Or tdoc = "ICA" Or tdoc = "ITR" Then ' Altas de Audit
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     ",det_ing_precio='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_ing_cantidad*det_ing_precio) as st " & _
                     " FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE ingreso " & _
                     " SET ing_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",ing_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' "
        ElseIf tdoc = "DCL" Then 'Devoluciones Cliente - Notas de Credito
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DPR" Then 'Devoluciones de Proyecto
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DRE" Then ' Devoluciones de Guias de Remision
            prec = CostoGuiaCli(prd_codigo, Doc)
            If exi + ing <> 0 Then
                PPP = FormatoD4((PPP * exi + ing * prec) / (exi + ing))
            End If
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     ",det_ing_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_ing_cantidad*det_ing_precio) as st " & _
                     " FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE ingreso " & _
                     " SET ing_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",ing_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' "
        ElseIf tdoc = "IRE" Then ' Devoluciones de Reservas
            prec = CostoReservaCli(prd_codigo, Doc)
            If exi + ing <> 0 Then
                PPP = FormatoD4((PPP * exi + ing * prec) / (exi + ing))
            End If
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "' " & _
                     ",det_ing_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_ing_cantidad*det_ing_precio) as st " & _
                     " FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE ingreso " & _
                     " SET ing_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",ing_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' "
        ElseIf tdoc = "IGR" Then ' Guias de Proveedor
            aux = PPP
            PPP = UbicaProximaAdqui(prd_codigo, docfecha, exi, PPP)
            prec = FormatoD4((PPP * (exi + ing) - aux * exi) / ing)
            strSql = " UPDATE det_ingreso " & _
                     " SET det_ing_costo='" & PPP & "', " & _
                     " det_ing_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_ing_cantidad*det_ing_precio) as st " & _
                     " FROM det_ingreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE ingreso " & _
                     " SET ing_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",ing_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_ing_codigo='" & tdoc & "' " & _
                     " AND ing_codigo='" & Doc & "' "
        End If
        CalculaPPP = PPP
        clsAux.Ejecutar strSql, "M"
    ElseIf d = 2 Then
        If tdoc = "FAC" Or tdoc = "NOT" Then  'Facturas
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "NET" Then ' Nota de Entrega Suministro
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "DPV" Then ' Devoluciones de Proveedor
            If FormatoD4(exi - egr) <> 0 Then
                PPP = FormatoD4((PPP * exi - egr * prec) / FormatoD4(exi - egr))
            End If
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "BAU" Or tdoc = "ECA" Or tdoc = "ETR" Then 'Bajas de Auditoria
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     ",det_egr_precio='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_egr_cantidad*det_egr_precio) as st " & _
                     " FROM det_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE egreso " & _
                     " SET egr_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",egr_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' "
        ElseIf tdoc = "ETN" Then 'EGRESO TRANFORMACION MATERIA PRIMA
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     ",det_egr_precio='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_egr_cantidad*det_egr_precio) as st " & _
                     " FROM det_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE egreso " & _
                     " SET egr_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",egr_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' "
        ElseIf tdoc = "NRP" Then ' Devolucion de Proyectos
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "GRE" Then ' Guias de Remision
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "ERE" Then ' Remision
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
        ElseIf tdoc = "EGR" Then 'Devoluciones de Guia de Proveedor
            prec = CostoGuiaPro(prd_codigo, Doc)
            If exi - egr <> 0 Then
                PPP = FormatoD4((PPP * exi - egr * prec) / (exi - egr))
            End If
            strSql = " UPDATE det_egreso " & _
                     " SET det_egr_costo='" & PPP & "' " & _
                     ",det_egr_precio='" & prec & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " AND prd_codigo='" & prd_codigo & "' "
            clsAux.Ejecutar strSql, "M"
            strSql = " SELECT SUM(det_egr_cantidad*det_egr_precio) as st " & _
                     " FROM det_egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' " & _
                     " GROUP BY emp_codigo "
            clsAux.Ejecutar strSql, "M"
            strSql = " UPDATE egreso " & _
                     " SET egr_subtotal='" & clsAux.adorec_Def("st") & "' " & _
                     ",egr_total='" & clsAux.adorec_Def("st") & "' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='" & tdoc & "' " & _
                     " AND egr_codigo='" & Doc & "' "
        End If
        CalculaPPP = PPP
        clsAux.Ejecutar strSql, "M"
    Else
        CalculaPPP = PPP
    End If
    Set clsAux = Nothing
End Function
Private Sub ActualizaCosto(prd_codigo As String, PPP As Double)
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " UPDATE producto " & _
             " SET prd_costo='" & PPP & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND prd_codigo='" & prd_codigo & "' "
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Sub
Private Function UbicaProximaAdqui(prd_codigo As String, doc_fecha As String, exi As Double, PPP As Double) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT TOP 1 det_ing_precio,det_ing_cantidad " & _
             " FROM ingreso INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " AND ingreso.ing_codigo=det_ingreso.ing_codigo" & _
             " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo" & _
             " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ing_fecha>='" & doc_fecha & "' " & _
             " AND (ingreso.tip_ing_codigo='COM' OR ingreso.tip_ing_codigo='IIM') " & _
             " AND ingreso.ing_anulado=0 " & _
             " ORDER BY ing_fecha "
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        PPP = FormatoD4((PPP * exi + clsAux.adorec_Def("det_ing_cantidad") * clsAux.adorec_Def("det_ing_precio")) / (exi + clsAux.adorec_Def("det_ing_cantidad")))
    End If
    UbicaProximaAdqui = PPP
    Set clsAux = Nothing
End Function
Private Function CostoGuiaCli(prd_codigo As String, Doc As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ing_factura " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND ingreso.ing_anulado=0 " & _
             " AND tip_ing_codigo='DRE'"
    clsAux.Ejecutar strSql
    strSql = " SELECT det_egr_costo " & _
             " FROM egreso INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
             " AND egreso.egr_codigo=det_egreso.egr_codigo" & _
             " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo" & _
             " AND det_egreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_codigo='" & clsAux.adorec_Def("ing_factura") & "'" & _
             " AND egreso.tip_egr_codigo='GRE' " & _
             " AND egreso.egr_anulado=0 "
    clsAux.Ejecutar strSql
    If clsAux.adorec_Def.RecordCount > 0 Then
        CostoGuiaCli = clsAux.adorec_Def("det_egr_costo")
    Else
        CostoGuiaCli = 0
    End If
    strSql = " UPDATE det_ingreso " & _
             " SET det_ing_precio='" & clsAux.adorec_Def("det_egr_costo") & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND tip_ing_codigo='DRE' " & _
             " AND prd_codigo='" & prd_codigo & "'"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Function
Private Function CostoReservaCli(prd_codigo As String, Doc As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT ing_factura " & _
             " FROM ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND ingreso.ing_anulado=0 " & _
             " AND tip_ing_codigo='IRE'"
    clsAux.Ejecutar strSql
    strSql = " SELECT det_egr_costo " & _
             " FROM egreso INNER JOIN det_egreso ON egreso.emp_codigo=det_egreso.emp_codigo " & _
             " AND egreso.egr_codigo=det_egreso.egr_codigo" & _
             " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo" & _
             " AND det_egreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
             " AND egreso.egr_codigo='" & clsAux.adorec_Def("ing_factura") & "'" & _
             " AND egreso.tip_egr_codigo='ERE' " & _
             " AND egreso.egr_anulado=0 "
    clsAux.Ejecutar strSql
    CostoReservaCli = clsAux.adorec_Def("det_egr_costo")
    strSql = " UPDATE det_ingreso " & _
             " SET det_ing_precio='" & clsAux.adorec_Def("det_egr_costo") & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ing_codigo='" & Doc & "'" & _
             " AND tip_ing_codigo='IRE' " & _
             " AND prd_codigo='" & prd_codigo & "'"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Function
Private Function CostoGuiaPro(prd_codigo As String, Doc As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    clsAux.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT egr_factura " & _
             " FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & Doc & "'" & _
             " AND egreso.egr_anulado=0 " & _
             " AND tip_egr_codigo='EGR'"
    clsAux.Ejecutar strSql
    strSql = " SELECT det_ing_precio " & _
             " FROM ingreso INNER JOIN det_ingreso ON ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " AND ingreso.ing_codigo=det_ingreso.ing_codigo" & _
             " AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo" & _
             " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
             " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
             " AND ingreso.ing_codigo='" & clsAux.adorec_Def("egr_factura") & "'" & _
             " AND ingreso.tip_ing_codigo='IGR' " & _
             " AND ingreso.ing_anulado=0 "
    clsAux.Ejecutar strSql
    CostoGuiaPro = clsAux.adorec_Def("det_ing_precio")
    strSql = " UPDATE det_egreso " & _
             " SET det_egr_precio='" & clsAux.adorec_Def("det_ing_precio") & "' " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND egr_codigo='" & Doc & "'" & _
             " AND tip_egr_codigo='EGR' " & _
             " AND prd_codigo='" & prd_codigo & "'"
    clsAux.Ejecutar strSql, "M"
    Set clsAux = Nothing
End Function
Private Function UltimoPPP(prd_codigo As String, desde As String) As Double
    Dim clsAux As New clsConsulta
    Dim strSql As String
    Dim ing_fecha As String
    Dim ing_costo As Double
    Dim egr_fecha As String
    Dim egr_costo As Double
    On Error Resume Next
    UltimoPPP = 0
    clsAux.Inicializar AdoConn, AdoConnMaster
'    strSql = " SELECT max(ing_fecha) as f " & _
'             " FROM ingreso INNER JOIN det_ingreso ON ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
'             " AND ingreso.emp_codigo=det_ingreso.emp_codigo " & _
'             " WHERE det_ing_costo!=0 " & _
'             " AND det_ingreso.emp_codigo='" & strEmpresa & "' AND ingreso.tip_ing_codigo!='ITR' " & _
'             " AND prd_codigo='" & prd_codigo & "' " & _
'             " AND ing_fecha < '" & desde & "' " & _
'             " AND ingreso.ing_anulado=0 " & _
'             " GROUP BY ingreso.emp_codigo"
    strSql = " SELECT TOP 1 ing_fecha as f " & _
             " FROM ingreso INNER JOIN det_ingreso ON ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
             " AND ingreso.emp_codigo=det_ingreso.emp_codigo " & _
             " WHERE det_ing_costo!=0 " & _
             " AND det_ingreso.emp_codigo='" & strEmpresa & "' AND ingreso.tip_ing_codigo!='ITR' " & _
             " AND prd_codigo='" & prd_codigo & "' " & _
             " AND ing_fecha < '" & desde & "' " & _
             " AND ingreso.ing_anulado=0 " & _
             " ORDER BY ing_fecha DESC"
    clsAux.Ejecutar strSql
    ing_fecha = clsAux.adorec_Def("f")
'    strSql = " SELECT max(egr_fecha) as f " & _
'             " FROM egreso INNER JOIN det_egreso " & _
'             " ON egreso.egr_codigo=det_egreso.egr_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
'             " AND egreso.emp_codigo=det_egreso.emp_codigo " & _
'             " WHERE det_egr_costo!=0 " & _
'             " AND egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo!='ETR' " & _
'             " AND prd_codigo='" & prd_codigo & "' " & _
'             " AND egr_fecha < '" & desde & "' " & _
'             " AND egreso.tip_egr_codigo!='EGR' AND egreso.tip_egr_codigo!='GRE' AND egreso.tip_egr_codigo!='ERE' " & _
'             " AND egreso.egr_anulado=0 " & _
'             " GROUP BY egreso.emp_codigo"
    strSql = " SELECT TOP 1 egr_fecha as f " & _
             " FROM egreso INNER JOIN det_egreso " & _
             " ON egreso.egr_codigo=det_egreso.egr_codigo AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
             " AND egreso.emp_codigo=det_egreso.emp_codigo " & _
             " WHERE det_egr_costo!=0 " & _
             " AND egreso.emp_codigo='" & strEmpresa & "' AND egreso.tip_egr_codigo!='ETR' " & _
             " AND prd_codigo='" & prd_codigo & "' " & _
             " AND egr_fecha < '" & desde & "' " & _
             " AND egreso.tip_egr_codigo!='EGR' AND egreso.tip_egr_codigo!='GRE' AND egreso.tip_egr_codigo!='ERE' " & _
             " AND egreso.egr_anulado=0 " & _
             " ORDER BY egr_fecha DESC"
    clsAux.Ejecutar strSql
    egr_fecha = clsAux.adorec_Def("f")
    If ing_fecha > egr_fecha Then
        strSql = " SELECT TOP 1 ing_fecha,det_ing_costo " & _
                 " FROM ingreso INNER JOIN det_ingreso ON ingreso.ing_codigo=det_ingreso.ing_codigo AND ingreso.tip_ing_codigo=det_ingreso.tip_ing_codigo " & _
                 " AND ingreso.emp_codigo=det_ingreso.emp_codigo " & _
                 " WHERE det_ing_costo!=0 " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & prd_codigo & "' " & _
                 " AND ing_fecha = '" & ing_fecha & "' " & _
                 " AND ingreso.ing_anulado=0 " & _
                 " ORDER BY ing_fecha DESC, ingreso.tip_ing_codigo DESC, ingreso.ing_codigo DESC"
        clsAux.Ejecutar strSql
        If clsAux.adorec_Def.RecordCount > 0 Then
            ing_fecha = clsAux.adorec_Def("ing_fecha")
            ing_costo = clsAux.adorec_Def("det_ing_costo")
        Else
            ing_costo = 0
        End If
        UltimoPPP = ing_costo
    End If
    If UltimoPPP <= 0 Then
        strSql = " SELECT TOP 1 egr_fecha,det_egr_costo " & _
                 " FROM egreso " & _
                 " INNER JOIN det_egreso ON egreso.egr_codigo=det_egreso.egr_codigo " & _
                 " AND egreso.tip_egr_codigo=det_egreso.tip_egr_codigo " & _
                 " AND egreso.emp_codigo=det_egreso.emp_codigo " & _
                 " WHERE det_egr_costo!=0 " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND prd_codigo='" & prd_codigo & "' " & _
                 " AND egr_fecha = '" & clsAux.adorec_Def("f") & "' " & _
                 " AND egreso.tip_egr_codigo NOT IN('EGR','GRE','ERE') " & _
                 " AND egreso.egr_anulado=0 " & _
                 " ORDER BY egr_fecha DESC, egreso.tip_egr_codigo DESC, egreso.egr_codigo DESC"
        clsAux.Ejecutar strSql
        If clsAux.adorec_Def.RecordCount > 0 Then
            egr_fecha = clsAux.adorec_Def("egr_fecha")
            egr_costo = clsAux.adorec_Def("det_egr_costo")
        Else
            egr_costo = 0
        End If
        UltimoPPP = egr_costo
    End If
    Set clsAux = Nothing
    If UltimoPPP < 0 Then
        UltimoPPP = 0
    End If
End Function
Public Sub CreaKardex(prd_codigo As String, desde As String, hasta As String)
    Dim strSql As String
    Dim strRangoFechaIng As String
    Dim strRangoFechaEgr As String
    If hasta = "" Then
        strRangoFechaEgr = "egreso.egr_fecha >= '" & desde & "' "
        strRangoFechaIng = "ingreso.ing_fecha >= '" & desde & "' "
    Else
        strRangoFechaEgr = "egreso.egr_fecha BETWEEN '" & desde & "' AND '" & hasta & "' "
        strRangoFechaIng = "ingreso.ing_fecha BETWEEN '" & desde & "' AND '" & hasta & "' "
    End If
    
    strSql = " EXEC Sp_RevisionExistencias"
    clsCon_Def.Ejecutar strSql, "M"

        strSql = " EXEC Sp_CreateTemporaryTable_If_Not_Exist '##ProdCosteado" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "', 'CREATE TABLE ##ProdCosteado" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "( " & _
                 " prd_codigo varchar(40) NOT NULL DEFAULT '''' " & _
                 " )' "
        clsCon_Def.Ejecutar strSql
        If prd_codigo = "%" Then
        strSql = " CREATE TABLE #ProdMovAUX( " & _
                 " prd_codigo varchar(40) NOT NULL DEFAULT '',emp_codigo CHAR(3) NOT NULL DEFAULT '' " & _
                 " ) "
                
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #ProdMovAUX " & _
                 " SELECT DISTINCT prd_codigo,egreso.emp_codigo " & _
                 " FROM (egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo = det_egreso.tip_egr_codigo " & _
                 " AND egreso.emp_codigo = det_egreso.emp_codigo AND egreso.egr_codigo = det_egreso.egr_codigo " & _
                 "  " & _
                 " AND det_egreso.prd_codigo LIKE '" & prd_codigo & "' ) " & _
                 " WHERE " & strRangoFechaEgr & _
                 " AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND egreso.egr_anulado=0 "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #ProdMovAUX " & _
                 " SELECT DISTINCT prd_codigo,ingreso.emp_codigo " & _
                 " FROM (ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo " & _
                 " AND ingreso.emp_codigo = det_ingreso.emp_codigo AND ingreso.ing_codigo = det_ingreso.ing_codigo " & _
                 "  " & _
                 " AND det_ingreso.prd_codigo LIKE '" & prd_codigo & "') " & _
                 " WHERE " & strRangoFechaIng & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                 " AND ingreso.ing_anulado=0 "
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TABLE #ProdMov (prd_codigo varchar(40) NOT NULL DEFAULT '',emp_codigo CHAR(3) NOT NULL DEFAULT '', PRIMARY KEY(prd_codigo,emp_codigo))"
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #ProdMov " & _
                 " SELECT DISTINCT RTRIM(LTRIM(prd_codigo)) as prd_codigo,emp_codigo " & _
                 " FROM #ProdMovAUX ProdMovAUX " & _
                 " WHERE  " & _
                 " prd_codigo LIKE '" & prd_codigo & "' " & _
                 " ORDER BY prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_Exist '#ProdMovAUX' "
        clsCon_Def.Ejecutar strSql
        
        strSql = " CREATE TABLE #AuxiExis ( " & _
                 " emp_codigo char(3) not null default '', " & _
                 " prd_codigo varchar(40) not null default '', " & _
                 " exi decimal(14,4) " & _
                 " ) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT egreso.emp_codigo, det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) " & _
                 " FROM (egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo = det_egreso.tip_egr_codigo " & _
                 " AND egreso.emp_codigo = det_egreso.emp_codigo AND egreso.egr_codigo = det_egreso.egr_codigo) " & _
                 " INNER JOIN #ProdMov ProdMov ON det_egreso.prd_codigo=ProdMov.prd_codigo AND  det_egreso.emp_codigo=ProdMov.emp_codigo " & _
                 " WHERE egreso.egr_fecha >= '" & desde & "' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND egreso.egr_anulado=0 " & _
                 " GROUP BY egreso.emp_codigo, det_egreso.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT ingreso.emp_codigo,det_ingreso.prd_codigo, SUM(det_ingreso.det_ing_cantidad) * (-1) " & _
                 " FROM (ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo " & _
                 " AND ingreso.emp_codigo = det_ingreso.emp_codigo AND ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                 " INNER JOIN #ProdMov ProdMov ON det_ingreso.prd_codigo=ProdMov.prd_codigo  AND  det_ingreso.emp_codigo=ProdMov.emp_codigo " & _
                 " WHERE ingreso.ing_fecha >= '" & desde & "' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND ingreso.ing_anulado=0 " & _
                 " GROUP BY ingreso.emp_codigo,det_ingreso.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT existencia.emp_codigo,existencia.prd_codigo, SUM(existencia.exi_cantidad) AS exi " & _
                 " FROM existencia " & _
                 " INNER JOIN #ProdMov ProdMov ON existencia.prd_codigo=ProdMov.prd_codigo AND existencia.emp_codigo=ProdMov.emp_codigo " & _
                 " WHERE existencia.emp_codigo='" & strEmpresa & "' " & _
                 " GROUP BY existencia.emp_codigo,existencia.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_Exist '#kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "'"
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TABLE #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " (  tdoc varchar(24) NOT NULL default '', doc decimal(15,0) NOT NULL default '0'," & _
                 " docfecha varchar(10) NOT NULL, " & _
                 " prd_codigo varchar(40) NOT NULL default '', " & _
                 " egr decimal(14,4) NOT NULL default '0', ing decimal(14,4) NOT NULL default '0',exi decimal(14,4) default NULL, " & _
                 " dep_codigo char(3) NOT NULL ,d decimal(1,0) NOT NULL default '0', " & _
                 " factura varchar(30) default NULL,prec decimal(14,4) NOT NULL default '0.0000', " & _
                 " PC decimal(14,4) default '0.0000' " & _
                 " ) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT 'SALDO INICIAL A LA FECHA' as tdoc, 0 as doc, '" & DateAdd("d", -1, desde) & "' as docfecha, producto.prd_codigo, 0 AS egr, " & _
                 " 0 AS ing, SUM(AuxiExis.exi) AS exi, '' as dep,0 as d, '' as fac,0.0000 as prec,prd_costo as PC " & _
                 " FROM (producto INNER JOIN #AuxiExis AuxiExis ON producto.prd_codigo = AuxiExis.prd_codigo AND producto.emp_codigo = AuxiExis.emp_codigo) " & _
                 " INNER JOIN #ProdMov ProdMov ON producto.prd_codigo=ProdMov.prd_codigo AND producto.emp_codigo=ProdMov.emp_codigo " & _
                 " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_codigo LIKE '" & prd_codigo & "'" & _
                 " GROUP BY producto.prd_codigo,prd_costo "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_EXist '#AuxiExis'"
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT ingreso.tip_ing_codigo as tdoc,ingreso.ing_codigo as doc,ingreso.ing_fecha as docfecha,det_ingreso.prd_codigo, 0 AS egr, " & _
                 " det_ingreso.det_ing_cantidad AS ing, 0 AS exi, det_ingreso.dep_codigo,1 as d, ingreso.ing_factura as factura,det_ing_precio,det_ing_costo " & _
                 " FROM ingreso INNER JOIN det_ingreso ON det_ingreso.ing_codigo=ingreso.ing_codigo AND det_ingreso.emp_codigo = ingreso.emp_codigo AND det_ingreso.tip_ing_codigo = ingreso.tip_ing_codigo AND det_ingreso.prd_codigo LIKE '" & prd_codigo & "' " & _
                 " INNER JOIN #ProdMov ProdMov ON det_ingreso.prd_codigo=ProdMov.prd_codigo AND det_ingreso.emp_codigo=ProdMov.emp_codigo " & _
                 " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND ingreso.ing_anulado=0 AND " & strRangoFechaIng & " " & _
                 " ORDER BY prd_codigo,tdoc,ingreso.ing_fecha "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT egreso.tip_egr_codigo as tdoc,egreso.egr_codigo as doc,egreso.egr_fecha as docfecha, det_egreso.prd_codigo, det_egreso.det_egr_cantidad AS egr, " & _
                 " 0 AS ing, 0 AS exi, det_egreso.dep_codigo,2 as d, egreso.egr_factura as factura,det_egr_precio,det_egr_costo " & _
                 " FROM egreso INNER JOIN det_egreso ON det_egreso.egr_codigo=egreso.egr_codigo AND det_egreso.emp_codigo = egreso.emp_codigo AND det_egreso.tip_egr_codigo = egreso.tip_egr_codigo AND det_egreso.prd_codigo LIKE '" & prd_codigo & "'" & _
                 " INNER JOIN #ProdMov ProdMov ON det_egreso.prd_codigo=ProdMov.prd_codigo AND det_egreso.emp_codigo=ProdMov.emp_codigo" & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND egreso.egr_anulado=0 AND " & strRangoFechaEgr & " " & _
                 " ORDER BY prd_codigo,tdoc,egreso.egr_fecha "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_Exist '#ProdMov'"
        clsCon_Def.Ejecutar strSql
    Else
        strSql = " CREATE TABLE #AuxiExis ( " & _
                 " emp_codigo char(3) not null default '', " & _
                 " prd_codigo varchar(40) not null default '', " & _
                 " exi decimal(14,4) " & _
                 " ) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT egreso.emp_codigo, det_egreso.prd_codigo, Sum(det_egreso.det_egr_cantidad) " & _
                 " FROM (egreso INNER JOIN det_egreso ON egreso.tip_egr_codigo = det_egreso.tip_egr_codigo " & _
                 " AND egreso.emp_codigo = det_egreso.emp_codigo AND egreso.egr_codigo = det_egreso.egr_codigo) " & _
                 " WHERE egreso.egr_fecha >= '" & desde & "' " & _
                 " AND egreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND det_egreso.prd_codigo='" & prd_codigo & "' " & _
                 " AND egreso.egr_anulado=0 " & _
                 " GROUP BY egreso.emp_codigo, det_egreso.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT ingreso.emp_codigo,det_ingreso.prd_codigo, SUM(det_ingreso.det_ing_cantidad) * (-1) " & _
                 " FROM (ingreso INNER JOIN det_ingreso ON ingreso.tip_ing_codigo = det_ingreso.tip_ing_codigo " & _
                 " AND ingreso.emp_codigo = det_ingreso.emp_codigo AND ingreso.ing_codigo = det_ingreso.ing_codigo) " & _
                 " WHERE ingreso.ing_fecha >= '" & desde & "' " & _
                 " AND ingreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
                 " AND ingreso.ing_anulado=0 " & _
                 " GROUP BY ingreso.emp_codigo,det_ingreso.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #AuxiExis " & _
                 " SELECT existencia.emp_codigo,existencia.prd_codigo, SUM(existencia.exi_cantidad) AS exi " & _
                 " FROM existencia " & _
                 " WHERE existencia.emp_codigo='" & strEmpresa & "' " & _
                 " AND existencia.prd_codigo='" & prd_codigo & "' " & _
                 " GROUP BY existencia.emp_codigo,existencia.prd_codigo "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_Exist '#kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & "'"
        clsCon_Def.Ejecutar strSql
        strSql = " CREATE TABLE #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " (  tdoc varchar(24) NOT NULL default '', doc decimal(15,0) NOT NULL default '0'," & _
                 " docfecha varchar(10) NOT NULL, " & _
                 " prd_codigo varchar(40) NOT NULL default '', " & _
                 " egr decimal(14,4) NOT NULL default '0', ing decimal(14,4) NOT NULL default '0',exi decimal(14,4) default NULL, " & _
                 " dep_codigo char(3) NOT NULL ,d decimal(1,0) NOT NULL default '0', " & _
                 " factura char(30) default NULL,prec decimal(14,4) NOT NULL default '0.0000', " & _
                 " PC decimal(14,4) default '0.0000' " & _
                 " ) "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT 'SALDO INICIAL A LA FECHA' as tdoc, 0 as doc, '" & DateAdd("d", -1, desde) & "' as docfecha, producto.prd_codigo, 0 AS egr, " & _
                 " 0 AS ing, SUM(AuxiExis.exi) AS exi, '' as dep,0 as d, '' as fac,0.0000 as prec,prd_costo as PC " & _
                 " FROM (producto INNER JOIN #AuxiExis AuxiExis ON producto.prd_codigo = AuxiExis.prd_codigo AND producto.emp_codigo = AuxiExis.emp_codigo) " & _
                 " WHERE producto.emp_codigo='" & strEmpresa & "' AND producto.prd_codigo ='" & prd_codigo & "'" & _
                 " GROUP BY producto.prd_codigo,prd_costo "
        clsCon_Def.Ejecutar strSql
        strSql = " EXEC Sp_Drop_Table_if_Exist '#AuxiExis' "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT ingreso.tip_ing_codigo as tdoc,ingreso.ing_codigo as doc,ingreso.ing_fecha as docfecha,det_ingreso.prd_codigo, 0 AS egr, " & _
                 " det_ingreso.det_ing_cantidad AS ing, 0 AS exi, det_ingreso.dep_codigo,1 as d, ingreso.ing_factura as factura,det_ing_precio,det_ing_costo " & _
                 " FROM ingreso INNER JOIN det_ingreso ON det_ingreso.ing_codigo=ingreso.ing_codigo AND det_ingreso.emp_codigo = ingreso.emp_codigo AND det_ingreso.tip_ing_codigo = ingreso.tip_ing_codigo AND det_ingreso.prd_codigo LIKE '" & prd_codigo & "' " & _
                 " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND det_ingreso.prd_codigo='" & prd_codigo & "' " & _
                 " AND ingreso.ing_anulado=0 AND " & strRangoFechaIng & " " & _
                 " ORDER BY prd_codigo,tdoc,ingreso.ing_fecha "
        clsCon_Def.Ejecutar strSql
        strSql = " INSERT INTO #kardex" & Replace(Replace(Replace(Replace(Replace(prd_codigo, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " " & _
                 " SELECT egreso.tip_egr_codigo as tdoc,egreso.egr_codigo as doc,egreso.egr_fecha as docfecha, det_egreso.prd_codigo, det_egreso.det_egr_cantidad AS egr, " & _
                 " 0 AS ing, 0 AS exi, det_egreso.dep_codigo,2 as d, egreso.egr_factura as factura,det_egr_precio,det_egr_costo " & _
                 " FROM egreso INNER JOIN det_egreso ON det_egreso.egr_codigo=egreso.egr_codigo AND det_egreso.emp_codigo = egreso.emp_codigo AND det_egreso.tip_egr_codigo = egreso.tip_egr_codigo AND det_egreso.prd_codigo LIKE '" & prd_codigo & "'" & _
                 " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                 "  " & _
                 " AND det_egreso.prd_codigo='" & prd_codigo & "' " & _
                 " AND egreso.egr_anulado=0 AND " & strRangoFechaEgr & " " & _
                 " ORDER BY prd_codigo,tdoc,egreso.egr_fecha "
        clsCon_Def.Ejecutar strSql
    
    End If

End Sub

Private Function CostoTransformacion(Doc As String) As Double
    Dim strSql As String
    Dim clsAux As New clsConsulta
    Dim clsAux2 As New clsConsulta
    Dim cosTotal As Double
    Dim cantTotal As Double
    Dim CosteoAux As New clsCostear
    CosteoAux.Inicializar AdoConn, AdoConnMaster
    clsAux.Inicializar AdoConn, AdoConnMaster
    clsAux2.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT emp_codigo,prd_codigo " & _
             " FROM det_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='ETN' " & _
             " AND egr_codigo='" & Doc & "' "
    clsAux.Ejecutar strSql
    While Not clsAux.adorec_Def.EOF
        strSql = "select prd_codigo FROM ##ProdCosteado" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " WHERE prd_codigo='" & clsAux.adorec_Def("prd_codigo") & "'"
        clsAux2.Ejecutar strSql
        If clsAux2.adorec_Def.RecordCount <= 0 Then
            strSql = "INSERT INTO ##ProdCosteado" & Replace(Replace(Replace(Replace(Replace(produc, " ", "_"), "-", "_"), ".", "_"), "%", "_"), "/", "_") & " VALUES('" & clsAux.adorec_Def("prd_codigo") & "')"
            clsAux2.Ejecutar strSql
            CosteoAux.Recostear Format(FD, "yyyy-mm-dd"), Format(fh, "yyyy-mm-dd"), clsAux.adorec_Def("prd_codigo"), True
        End If
        clsAux.adorec_Def.MoveNext
    Wend
    strSql = " SELECT emp_codigo,SUM(det_egr_cantidad*det_egr_precio) as st " & _
             " FROM det_egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='ETN' " & _
             " AND egr_codigo='" & Doc & "' " & _
             " GROUP BY emp_codigo "
    clsAux.Ejecutar strSql
    cosTotal = clsAux.adorec_Def("st")
    strSql = " SELECT emp_codigo,SUM(det_ing_cantidad) as st " & _
             " FROM det_ingreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_ing_codigo='ITN' " & _
             " AND ing_codigo='" & Doc & "' " & _
             " GROUP BY emp_codigo "
    clsAux.Ejecutar strSql
    cantTotal = clsAux.adorec_Def("st")
    CostoTransformacion = cosTotal / cantTotal
End Function


Attribute VB_Name = "modconexion"
Public AdoConn As ADODB.Connection

Public GeneraDocElec As Integer
Public PtoEmiDocEle As String
Public ReservaNotaCredito As Integer

Public s_userSql As String
Public s_passwordSql As String
Public s_catalogoSQL As String
Public s_instanciaSQL As String
Public cadena_conexion As String

Public booUnContenedor As Boolean
Public strContenedorRecurrente As String

Public NombreComercial As String
Public CorreoServicioAlCliente As String
Public CorreoCartera As String
Public CorreoCompras As String
Public CorreoAsistenteCos As String
Public CorreoSupervisorDeTransportes As String
Public CorreoNoticias As String
Public CorreoAsistenteCartera As String

Public NoCopiasFactura As Integer
Public ContenedorInventario As Boolean

Public strServidorBDDLocal As String
Public strPuertoLocal As String
Public AdoConnMaster As ADODB.Connection
Public strServidorBDDMaster As String
Public strPuertoMaster As String
Public strBDD As String
Public strServidorWeb As String
Public strUsuario As String
Public strClave As String
Public strEmpresa As String
Public strSucursal As String
Public strSucursal2 As String
Public strBodega As String

Public ImpresoraEtiqueta As String
Public ImpresoraTicket As String
Public ImpresoraPorDefecto As String
Public ImprimirEtiquetaDespacho As Boolean

Public PuertoBalanza As Integer

Public strPtoFactura As String
Public strPtoFacturaOriginal As String
Public strAutorFactura As String
Public strCaducaFactura As String
Public strCodCuenta As String, strDescCuenta As String, strEstadoPYG As String
Public lngFacturaDesde As Long
Public lngFacturaHasta As Long
Public intNivCta As Integer

Public PorIVA As Double
Public CodigoIVA As Integer

Public strPathHuella As String

Public Path As String
Public PathCpp As String
Public num As Integer
Public Archivo As String
Public archivoaux As String
Public nLen As Long
Public Texto As String
Public ATok As Integer
Public enroll As Double
Public hProcess As Long
Public strFecha As Date
Public Historico As Boolean

Public strForNumDec As String
Public strForNumMil As String
Public strForMonDec As String
Public strForMonMil As String
Public strForFecha As String
Public HoyDia As String
Public HoyDiaHora As String

Public strServidorSMTP As String
Public strUsuarioSMTP As String
Public strClaveSMTP As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long

Public Const SND_FILENAME = &H20000


Public Function FormatoFecha(Fecha As Variant)
    FormatoFecha = Format(Fecha, "yyyy-MM-dd")
End Function

Public Function FormatoHora(Hora As Variant)
    FormatoHora = Format(Hora, "HH:mm")
End Function


Public Sub Seleccionar_Contenido()
    Screen.ActiveControl.SelStart = 0
    If UCase(TypeName(Screen.ActiveControl)) = "MASKEDBOX" Then
        Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.FormattedText)
    Else
        Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
    End If
End Sub

Public Function LeerArchivo(Archivo As String, LineaInicio As Integer, Tam As Boolean, StrSepara As String, strFin As String) As Variant()
    'Javo
    Dim varObjeto() As Variant
    Dim strObjeto() As String
    
    Dim i As Integer
    Dim j As Integer
    Dim strPath As String
    Dim strLinea As String
    Dim PosicionActual As Integer
    Dim vacio As Boolean
    Dim numLinea As Integer
    Dim st As Long
    Dim salir As Boolean
    

    numLinea = 0
    i = 0
    vacio = True
    If Dir(Archivo) <> "" Then
        Open Archivo For Input As #1
            While EOF(1) = False
                numLinea = numLinea + 1
                If numLinea >= LineaInicio Then
                    vacio = False
                    'Va agrandando el array
                    ReDim Preserve varObjeto(i + 1)
                    Input #1, strLinea
                    PosicionActual = 1
                    salir = False
                    j = 0
                    While salir = False
                        ReDim Preserve strObjeto(j + 1)
                        If Tam = False Then
                            st = InStr(PosicionActual, strLinea, StrSepara) - 1
                            If st < 0 Then
                                st = Len(strLinea) + 1
                                salir = True
                            End If
                        Else
                            st = FormatoD0(StrSepara)
                        End If
                        'Pone el texto en el objeto
                        If st - PosicionActual + 1 > 0 Then
                            strObjeto(j) = Mid(strLinea, PosicionActual, st - PosicionActual + 1)
                        End If
                        PosicionActual = st + 2
                        
                        j = j + 1
                    Wend
                    varObjeto(i) = strObjeto
                    i = i + 1
                Else
                    Input #1, strLinea
                End If
            Wend
        Close #1
        If vacio = True Then GoTo CasoVacio
    Else
        MsgBox "No se encuentra el archivo " & Archivo & ".", vbInformation
CasoVacio:
        'Devuelva una matriz con un array de cadena vacía
        ReDim varObjeto(0)
        ReDim strObjeto(0)
        LeerArchivo = varObjeto
        Exit Function
    End If
    
    LeerArchivo = varObjeto
End Function

Public Function FormatoD2(Numero As Variant) As Variant
    FormatoD2 = Val(Format(Numero, "###0.00"))
End Function

Public Function FormatoD4(Numero As Variant) As Variant
    FormatoD4 = Val(Format(Numero, "###0.0000"))
End Function

Public Function FormatoD8(Numero As Variant) As Variant
    FormatoD8 = Val(Format(Numero, "###0.00000000"))
End Function

Public Function FormatoD0(Numero As Variant) As Variant
    Dim x As Double
    x = Val(Format(Numero, "###0"))
    FormatoD0 = Val(Format(Numero, "###0.00"))
End Function


Public Function Lpad(Valor$, CaracterAumentar$, largo%) As String
    Dim x As Integer
    Dim PadLength As Integer
    Valor = Left(Valor, largo)
    PadLength = largo - Len(Valor)
    Dim PadString As String
    For x = 1 To PadLength
        PadString = PadString & CaracterAumentar
    Next
    Lpad = PadString + Valor

End Function

Public Function Rpad(Valor, CaracterAumentar$, largo%) As String
    Dim x As Integer
    Dim PadLength As Integer
    Valor = Left(Valor, largo)
    PadLength = largo - Len(Valor)
    Dim PadString As String
    For x = 1 To PadLength
        PadString = CaracterAumentar & PadString
    Next
    Rpad = Valor & PadString
   End Function


Public Function Bool(cadena As Variant)
    If cadena = "" Then
        Bool = False
    Else
        Bool = CBool(Val(cadena))
    End If
End Function
'Comienzan funciones que definen el módulo de recursos humanos

Function GrabarDescuento(clsDescuento As clsConsulta, Tipo As String, Empleado As String, Fecha As String, Valor As Double, Optional Asiento As String)
    'Dim clsDescuento As New clsConsulta
    Dim codigo As Long
    Dim strFecha As String
    Dim strSql As String
    Dim strAsiento As String
    'clsDescuento.Inicializar AdoConn
Repetir:
    strSql = " SELECT IFNULL(MAX(des_codigo),0)+1 AS num " & _
                  " FROM descuento" & _
                  " WHERE emp_codigo='" & strEmpresa & "' "
    clsDescuento.Ejecutar (strSql)
    codigo = clsDescuento.adorec_Def("num")
    If Trim(Asiento) <> "" Then
        strAsiento = "'" & Asiento & "'"
    Else
        strAsiento = "NULL"
    End If
    strSql = " INSERT INTO descuento " & _
                  " ( des_codigo, emp_codigo, tip_des_codigo,  " & _
                  "   epl_codigo, des_fecha, des_valor, des_pagado, asi_numasiento, des_fechamod, des_usumod ) " & _
                  " VALUES " & _
                  " ('" & codigo & "', '" & strEmpresa & "', '" & Tipo & "'," & _
                  " '" & Empleado & "', '" & Fecha & "', '" & Valor & "'," & _
                  " 0, " & strAsiento & ",CURRENT_TIMESTAMP, '" & strUsuario & "') "
    clsDescuento.Ejecutar strSql
    'GrabarDescuento = Codigo
End Function

Sub ActualizarDescuento(codigo As String, Cliente As String, capital As Double, interes As Double, Optional Valor1 As String, Optional Valor2 As String, Optional Asiento As String, Optional GrabarMovimiento As Boolean)
    Dim clsDescuento As New clsConsulta
    Dim strSql As String
    Dim strSet As String
    Dim Condicion1 As String
    Dim Condicion2 As String
    Dim Condicion3 As String
    clsDescuento.Inicializar AdoConn, AdoConnMaster
    If Valor1 <> "" Then
        Condicion1 = " des_valor1='" & FormatoD2(Valor1) & "',"
    End If
    If Valor2 <> "" Then
        Condicion2 = " des_valor2='" & FormatoD2(Valor2) & "',"
    End If
    If Asiento <> "" Then
        Condicion3 = " asi_numasiento='" & Asiento & "',"
    End If
    If GrabarMovimiento = False Then
        strSet = "epl_codigo='" & Cliente & "',"
    Else
        strSet = "tip_des_codigo='" & Cliente & "',"
    End If
    strSql = " UPDATE descuento SET " & strSet & _
             " des_valor='" & capital & "'," & Condicion1 & Condicion2 & Condicion3 & _
             " des_fechamod=CURRENT_TIMESTAMP, des_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' AND des_codigo='" & codigo & "'"
    clsDescuento.Ejecutar strSql
End Sub

Sub EliminarDescuento(clsDescuento As clsConsulta, codigo As Long, Fecha As String)
    'Dim clsDescuento As New clsConsulta
    Dim strSql As String
    'clsDescuento.Inicializar AdoConn
    
    strSql = " DELETE FROM descuento " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND epl_codigo='" & codigo & "' " & _
             " AND des_fecha='" & Fecha & "' "
    clsDescuento.Ejecutar strSql
End Sub

Function SueldoMes(Empleado As String, Fecha1 As Variant, Fecha2 As Variant) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT IFNULL(SUM(CASE tip_des_ingreso WHEN 0 THEN des_valor*-1 ELSE des_valor END),0) " & _
             " FROM descuento INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_sueldo_mes=1 AND epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        SueldoMes = clsSql1.adorec_Def(0)
    Else
        SueldoMes = 0
    End If
End Function

Function SueldoAño(Empleado As String, Fecha1 As Variant, Fecha2 As Variant) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT IFNULL(SUM(CASE tip_des_ingreso WHEN 0 THEN des_valor*-1 ELSE des_valor END),0) " & _
             " FROM descuento INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Left(Fecha1, 4) & "-01-01' AND '" & Left(Fecha1, 4) & "-12-31'" & _
             " AND tip_des_sueldo_mes=1 AND epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        SueldoAño = clsSql1.adorec_Def(0)
    Else
        SueldoAño = 0
    End If
End Function

Function SueldoIESS(Empleado As String, Fecha1 As Variant, Fecha2 As Variant) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT IFNULL(SUM(CASE tip_des_ingreso WHEN 0 THEN des_valor*-1 ELSE des_valor END),0) " & _
             " FROM descuento INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_iess=1 AND epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        SueldoIESS = clsSql1.adorec_Def(0)
    Else
        SueldoIESS = 0
    End If
End Function



Function RentaMes(Empleado As String, Fecha1 As Variant, Fecha2 As Variant) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    'Sumar los valores del mes de los tipos que esten marcados para el calculo del impuesto
    strSql = " SELECT IFNULL(SUM(CASE tip_des_ingreso WHEN 0 THEN des_valor*-1 ELSE des_valor END),0) " & _
             " FROM descuento INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_impuesto_renta=1 AND epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        RentaMes = clsSql1.adorec_Def(0)
    Else
        RentaMes = 0
    End If
End Function

Function ImpuestoRentaMes(Valor As Double) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    'Sumar los valores del mes de los tipos que esten marcados para el calculo del impuesto
    strSql = " SELECT ren_frac_basica, ren_imp_frac_basica, ren_imp_frac_excedente FROM parametro_renta" & _
             " WHERE ren_frac_basica <= " & Valor & " AND ren_frac_exceso >= " & Valor & _
             " AND ren_codigo LIKE 'M%' AND emp_codigo='" & strEmpresa & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        ImpuestoRentaMes = clsSql1.adorec_Def("ren_imp_frac_basica")
        ImpuestoRentaMes = ImpuestoRentaMes + FormatoD2((Valor - Val(clsSql1.adorec_Def("ren_frac_basica"))) * Val(clsSql1.adorec_Def(2)) / 100)
    Else
        ImpuestoRentaMes = 0
    End If
End Function

Function RentaAño(Empleado As String, Fecha1 As Variant, Fecha2 As Variant) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    'Sumar los valores del mes de los tipos que esten marcados para el calculo del impuesto
    strSql = " SELECT IFNULL(SUM(CASE tip_des_ingreso WHEN 0 THEN des_valor*-1 ELSE des_valor END),0) " & _
             " FROM descuento INNER JOIN tipo_descuento ON descuento.tip_des_codigo=tipo_descuento.tip_des_codigo AND descuento.emp_codigo=tipo_descuento.emp_codigo" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "' AND des_fecha BETWEEN '" & Fecha1 & "' AND '" & Fecha2 & "'" & _
             " AND tip_des_impuesto_renta=1 AND epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        RentaAño = clsSql1.adorec_Def(0)
    Else
        RentaAño = 0
    End If
End Function

Function ImpuestoRentaAño(Valor As Double) As Double
    Dim strSql As String
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    'Sumar los valores del mes de los tipos que esten marcados para el calculo del impuesto
    strSql = " SELECT ren_frac_basica, ren_imp_frac_basica, ren_imp_frac_excedente FROM parametro_renta" & _
             " WHERE ren_frac_basica <= " & Valor & " AND ren_frac_exceso >= " & Valor & _
             " AND ren_codigo LIKE 'A%' AND emp_codigo='" & strEmpresa & "'"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        ImpuestoRentaAño = clsSql1.adorec_Def("ren_imp_frac_basica")
        ImpuestoRentaAño = ImpuestoRentaAño + FormatoD2((Valor - clsSql1.adorec_Def("ren_frac_basica")) * Val(clsSql1.adorec_Def("ren_imp_frac_excedente")) / 100)
    Else
        ImpuestoRentaAño = 0
    End If
End Function

Function SumarProvisionesPendientes(codigo As String, Empleado As String, Provisionada As String) As Double
    Dim strSql As String
    Dim strWhere As String
    Dim PTotal As Double
    Dim Pparcial As Double
    Dim clsSql1 As New clsConsulta
    clsSql1.Inicializar AdoConn, AdoConnMaster
    'Buscar fecha del mes en el que hubo provisión y entrega de provisión
    strSql = " SELECT descuento.des_fecha FROM descuento" & _
             " INNER JOIN descuento des2 ON descuento.emp_codigo=des2.emp_codigo AND descuento.des_fecha=des2.des_fecha AND descuento.epl_codigo=des2.epl_codigo AND des2.tip_des_codigo='" & Provisionada & "'" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "'" & _
             " AND descuento.tip_des_codigo='" & codigo & "' AND descuento.epl_codigo='" & Empleado & "'" & _
             " ORDER BY descuento.des_fecha DESC"
    clsSql1.Ejecutar (strSql)
    If clsSql1.adorec_Def.RecordCount > 0 Then
        strWhere = " AND des_fecha>" & clsSql1.adorec_Def(0) & " "
    Else
        strWhere = ""
    End If
    'Sumar el valor de las provisiones no entregadas
    strSql = " SELECT IFNULL(SUM(des_valor),0) AS valor FROM descuento" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "'" & strWhere & _
             " AND descuento.tip_des_codigo='" & codigo & "' AND descuento.epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    PTotal = CDbl(clsSql1.adorec_Def("valor"))
    strSql = " SELECT IFNULL(SUM(des_valor),0) AS valor FROM descuento" & _
             " WHERE descuento.emp_codigo='" & strEmpresa & "'" & strWhere & _
             " AND descuento.tip_des_codigo='" & Provisionada & "' AND descuento.epl_codigo='" & Empleado & "'"
    clsSql1.Ejecutar (strSql)
    Pparcial = CDbl(clsSql1.adorec_Def("valor"))
    SumarProvisionesPendientes = PTotal - Pparcial
End Function

Function DiasFinDeMes(FinDeMes As String, FechaIngreso As String, FechaSalida As String) As Integer
    If Trim(FinDeMes) = "" Or Trim(FechaIngreso) = "" Then
        DiasFinDeMes = 0
        Exit Function
    End If
    
    Dim dia As Integer
    Dim Mes As Integer
    Dim Año As Integer
    
    Dim DiaI As Integer
    Dim MesI As Integer
    Dim AñoI As Integer
    
    Dim dias As Integer
    Dim MesS As Integer
    Dim AñoS As Integer
    
    Dim IngresóMes As Boolean
    Dim SalióMes As Boolean
    
    dia = CInt(Mid(FinDeMes, 9, 2))
    Mes = CInt(Mid(FinDeMes, 6, 2))
    Año = CInt(Left(FinDeMes, 4))

    DiaI = CInt(Mid(FechaIngreso, 9, 2))
    MesI = CInt(Mid(FechaIngreso, 6, 2))
    AñoI = CInt(Left(FechaIngreso, 4))
    
    'Si el empleado entró este mes a trabajar tiene menos días
    If Mes = MesI And Año = AñoI Then
        IngresóMes = True
    End If
    
    If Trim(FechaSalida) <> "" Then
        dias = CInt(Mid(FechaSalida, 9, 2))
        MesS = CInt(Mid(FechaSalida, 6, 2))
        AñoS = CInt(Left(FechaSalida, 4))
        'Si el empleado salió este mes de trabajar tiene menos días
        If Mes = MesS And Año = AñoS Then
            SalióMes = True
        End If
    End If
    'Ingresó este mes
    If IngresóMes = True And SalióMes = False Then
        DiasFinDeMes = dia - DiaI + 1
    'Salió este mes
    ElseIf IngresóMes = False And SalióMes = True Then
        DiasFinDeMes = dias
    'Ingresó y salió este mes
    ElseIf IngresóMes = True And SalióMes = True Then
        DiasFinDeMes = dias - DiaI + 1
    Else
        'Para que salgan cero días cuando aún no ha entrado a la empresa
        If DateDiff("d", FechaIngreso, FinDeMes) > 0 Then
            DiasFinDeMes = dia
        Else
            DiasFinDeMes = 0
        End If
        'Que salgan cero días meses luego de haber salido
        If Trim(FechaSalida) <> "" Then
            If DateDiff("d", FechaSalida, FinDeMes) > 0 Then
                DiasFinDeMes = 0
            End If
        End If
    End If
End Function

Function DiasFondo(FinDeMes As String, FechaIngreso As String, FechaSalida As String) As Integer
    If Trim(FinDeMes) = "" Or Trim(FechaIngreso) = "" Then
        DiasFondo = 0
        Exit Function
    End If
    
    Dim dia As Integer
    Dim Mes As Integer
    Dim Año As Integer
    
    Dim DiaI As Integer
    Dim MesI As Integer
    Dim AñoI As Integer
    
    Dim dias As Integer
    Dim MesS As Integer
    Dim AñoS As Integer
    
    Dim CumpleañosMes As Boolean
    Dim SalióMes As Boolean
    
    dia = CInt(Mid(FinDeMes, 9, 2))
    Mes = CInt(Mid(FinDeMes, 6, 2))
    Año = CInt(Left(FinDeMes, 4))

    DiaI = CInt(Mid(FechaIngreso, 9, 2))
    MesI = CInt(Mid(FechaIngreso, 6, 2))
    AñoI = CInt(Left(FechaIngreso, 4))
    
    'Si el empleado justo cumple un año en este mes
    If Mes = MesI And Año = AñoI + 1 Then
        CumpleañosMes = True
    End If
    
    If Trim(FechaSalida) <> "" Then
        dias = CInt(Mid(FechaSalida, 9, 2))
        MesS = CInt(Mid(FechaSalida, 6, 2))
        AñoS = CInt(Left(FechaSalida, 4))
        'Si el empleado salió este mes de trabajar tiene menos días
        If Mes = MesS And Año = AñoS Then
            SalióMes = True
        End If
    End If
    'Para que salgan cero días cuando aún no tiene su cumpleaños número 1
    If DateDiff("d", DateAdd("yyyy", 1, FechaIngreso), FinDeMes) > 0 Then
        'Cumple años este mes
        If CumpleañosMes = True And SalióMes = False Then
            DiasFondo = dia - DiaI + 1
        'Salió este mes
        ElseIf CumpleañosMes = False And SalióMes = True Then
            DiasFondo = dias
        'Ingresó y salió este mes
        ElseIf CumpleañosMes = True And SalióMes = True Then
            DiasFondo = dias - DiaI + 1
        Else
            DiasFondo = dia
            'Que salgan cero días meses luego de haber salido
            If Trim(FechaSalida) <> "" Then
                If DateDiff("d", FechaSalida, FinDeMes) > 0 Then
                    DiasFondo = 0
                End If
            End If
        End If
    Else
        DiasFondo = 0
    End If
    
End Function

Sub NuevoDetalleActivo(CodigoActivo As String, Tipo As String, Asiento As String, Mes As Integer, Año As Integer, Valor As String)
    Dim clsSql As New clsConsulta
    Dim sentencia_SQL As String
    Dim maximo As String
    
    clsSql.Inicializar AdoConn, AdoConnMaster

    sentencia_SQL = " SELECT IFNULL(MAX(det_act_fij_codigo),0)+1 as num " & _
                    " FROM det_activo_fijo " & _
                    " WHERE emp_codigo='" & strEmpresa & "' " & _
                    " AND act_fij_codigo = '" & CodigoActivo & "'"
    clsSql.Ejecutar (sentencia_SQL)
    maximo = clsSql.adorec_Def(0)
    
        
    sentencia_SQL = " INSERT INTO det_activo_fijo (det_act_fij_codigo, act_fij_codigo, emp_codigo, det_act_fij_tipo, asi_numasiento," & _
                    " det_act_fij_mes, det_act_fij_año, det_act_fij_valor, det_act_fij_fechamod, det_act_fij_usumod)  " & _
                    " VALUES ('" & maximo & "','" & CodigoActivo & "', " & _
                    " '" & strEmpresa & "', '" & Tipo & "', " & _
                    " '" & Asiento & "','" & Mes & "', " & _
                    " " & Año & ", '" & FormatoD2(Valor) & "',CURRENT_TIMESTAMP, '" & strUsuario & "')"
    clsSql.Ejecutar (sentencia_SQL)
End Sub


Sub SeleccionarFlexGrid2(objeto As Object)
    On Error Resume Next 'El setfocus siempre causa problemas mejor prevenir
    objeto.SetFocus
    objeto.Select 0, 0, objeto.Rows - 1, objeto.Cols - 1
End Sub


Sub CopiarFlexGrid2(objetoVSFG As Object)
    Screen.MousePointer = vbHourglass
    Dim Copiar As String
    Dim x1, x2, y1, y2 As Long
    Dim PrimeraVez As Boolean
    If objetoVSFG.Row > objetoVSFG.RowSel Then
        x1 = objetoVSFG.RowSel
        x2 = objetoVSFG.Row
    Else
        x1 = objetoVSFG.Row
        x2 = objetoVSFG.RowSel
    End If
    If objetoVSFG.Col > objetoVSFG.ColSel Then
        y1 = objetoVSFG.ColSel
        y2 = objetoVSFG.Col
    Else
        y1 = objetoVSFG.Col
        y2 = objetoVSFG.ColSel
    End If
    Copiar = ""
    'MsgBox x1 & "-" & y1 & "    " & x2 & "-" & y2
    For xx = x1 To x2
        PrimeraVez = True
        For yy = y1 To y2
            If objetoVSFG.ColHidden(yy) = False Then
                If PrimeraVez = False Then Copiar = Copiar & vbTab
                Copiar = Copiar & objetoVSFG.Cell(flexcpTextDisplay, xx, yy)
                PrimeraVez = False
            End If
        Next yy
        Copiar = Copiar & vbNewLine
    Next xx
    'MsgBox Copiar
    Clipboard.Clear
    Clipboard.SetText Copiar
    Screen.MousePointer = vbDefault
End Sub

Public Function BorrarAplicacion() As Boolean
    Dim fso As New FileSystemObject
    On Error GoTo problema
    If fso.FolderExists(App.Path) = True Then
        fso.DeleteFolder App.Path, False
    End If
    Set fso = Nothing
    BorrarAplicacion = True
    Exit Function
problema:
    BorrarAplicacion = False
End Function


Public Function MostrarDia(Index As Integer) As String
    If Index = 0 Then
        MostrarDia = "Lunes"
    ElseIf Index = 1 Then
        MostrarDia = "Martes"
    ElseIf Index = 2 Then
        MostrarDia = "Miércoles"
    ElseIf Index = 3 Then
        MostrarDia = "Jueves"
    ElseIf Index = 4 Then
        MostrarDia = "Viernes"
    ElseIf Index = 5 Then
        MostrarDia = "Sábado"
    ElseIf Index = 6 Then
        MostrarDia = "Domingo"
    End If
End Function

Public Function MostrarMes(Index As Integer) As String
    If Index = 1 Then
        MostrarMes = "Enero"
    ElseIf Index = 2 Then
        MostrarMes = "Febrero"
    ElseIf Index = 3 Then
        MostrarMes = "Marzo"
    ElseIf Index = 4 Then
        MostrarMes = "Abril"
    ElseIf Index = 5 Then
        MostrarMes = "Mayo"
    ElseIf Index = 6 Then
        MostrarMes = "Junio"
    ElseIf Index = 7 Then
        MostrarMes = "Julio"
    ElseIf Index = 8 Then
        MostrarMes = "Agosto"
    ElseIf Index = 9 Then
        MostrarMes = "Septiembre"
    ElseIf Index = 10 Then
        MostrarMes = "Octubre"
    ElseIf Index = 11 Then
        MostrarMes = "Noviembre"
    ElseIf Index = 12 Then
        MostrarMes = "Diciembre"
    End If
End Function


Public Function MesCerrado(Fe As String) As Boolean
    Dim clsMesCerrado As New clsConsulta
    clsMesCerrado.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT COALESCE(COUNT(*),0) " & _
             " FROM cierre_mes " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cie_mes_ano=" & Year(Fe) & " AND cie_mes_mes=" & Month(Fe)
    clsMesCerrado.Ejecutar strSql
    MesCerrado = False
    If clsMesCerrado.adorec_Def.RecordCount > 0 Then
        If FormatoD0(clsMesCerrado.adorec_Def(0)) > 0 Then
            MsgBox "El mes está cerrado por contabilidad", vbInformation, "Fecha"
            MesCerrado = True
        End If
    End If

End Function

Public Sub IngresoBlocFactura()
    Do
        lngFacturaDesde = Abs(Val(InputBox("El bloc para el Negocio empieza en el número", "Numeración de Documento", lngFacturaDesde)))
    Loop While lngFacturaDesde = 0
    Do
        lngFacturaHasta = Abs(Val(InputBox("El bloc para el Negocio termina en el número", "Numeración de Documento", lngFacturaHasta)))
    Loop While lngFacturaHasta = 0
End Sub


Public Function code128$(chaine$)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Parameters : a string
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  code128$ = ""
  If Len(chaine$) > 0 Then
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculation of the code string with optimized use of tables B and C
    code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'See if interesting to switch to table C
          'yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choice of table C
            If i% = 1 Then 'Starting with table C
              code128$ = Chr$(210)
            Else 'Switch to table C
              code128$ = code128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then code128$ = Chr$(209) 'Starting with table B
          End If
        End If
        If Not tableB Then
          'We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            code128$ = code128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'We haven't 2 digits, switch to table B
            code128$ = code128$ & Chr$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Process 1 digit with table B
          code128$ = code128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calculation of the checksum
      For i% = 1 To Len(code128$)
        dummy% = Asc(Mid$(code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Add the checksum and the STOP
      code128$ = code128$ & Chr$(checksum&) & Chr$(211)
    End If
  End If
  Exit Function
testnum:
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function


Public Sub EnviarMail(strDeNombre As String, strDe As String, strParaNombre As String, strPara As String, strConCopiaOculta As String, strAsunto As String, strCuerpo As String, Optional strAdjunto As String = "", Optional booComoHTML As Boolean = False)
    Dim clsSMTP As New clsConsulta
    Dim strSql As String
    strDe = Replace(strDe, " ", "")
    strPara = Replace(strPara, " ", "")
    strConCopiaOculta = Replace(strConCopiaOculta, " ", "")
    If Left(Trim(strPara), 1) = ";" Then
        strPara = Trim(Mid(strPara, 2))
    End If
    If Right(Trim(strPara), 1) = ";" Then
        strPara = Trim(Left(strPara, Len(strPara) - 1))
    End If
    strConCopiaOculta = Trim(strConCopiaOculta)
    If Left(Trim(strConCopiaOculta), 1) = ";" Then
        strConCopiaOculta = Trim(Mid(strConCopiaOculta, 2))
    End If
    If Right(Trim(strConCopiaOculta), 1) = ";" Then
        strConCopiaOculta = Trim(Left(strConCopiaOculta, Len(strConCopiaOculta) - 1))
    End If
    If strServidorSMTP = "" Then
        clsSMTP.Inicializar AdoConn, AdoConnMaster
        strSql = "SELECT COALESCE(par_texto,'') as par_texto FROM parametro WHERE par_codigo='SMT' AND emp_codigo='RYB' "
        clsSMTP.Ejecutar strSql
        If clsSMTP.adorec_Def.RecordCount > 0 Then
            strServidorSMTP = clsSMTP.adorec_Def("par_texto")
        End If
        strSql = "SELECT COALESCE(par_texto,'') as par_texto FROM parametro WHERE par_codigo='SMU' AND emp_codigo='RYB' "
        clsSMTP.Ejecutar strSql
        If clsSMTP.adorec_Def.RecordCount > 0 Then
            strUsuarioSMTP = clsSMTP.adorec_Def("par_texto")
        End If
        strSql = "SELECT COALESCE(par_texto,'') as par_texto FROM parametro WHERE par_codigo='SMC' AND emp_codigo='RYB' "
        clsSMTP.Ejecutar strSql
        If clsSMTP.adorec_Def.RecordCount > 0 Then
            strClaveSMTP = clsSMTP.adorec_Def("par_texto")
        End If
    End If

    If strPara <> "" Then
        Dim FREmail As New frmEmail
        FREmail.Show
        FREmail.txtFromName = strDeNombre
        FREmail.txtFrom = strDe
        FREmail.txtToName.Text = strParaNombre
        FREmail.txtTo.Text = strPara
        FREmail.txtBcc.Text = strConCopiaOculta
        If booComoHTML = True Then
            FREmail.ckHtml.Value = 1
        Else
            FREmail.ckHtml.Value = 0
        End If
        FREmail.txtSubject.Text = Trim(strAsunto)
        FREmail.txtMsg.Text = strCuerpo
        FREmail.txtAttach.Text = strAdjunto
        FREmail.txtServer = strServidorSMTP
        FREmail.txtUserName = strUsuarioSMTP
        FREmail.txtPassword = strClaveSMTP
        If strUsuarioSMTP <> "" Then
            FREmail.ckLogin = 1
        End If
        FREmail.cmdSend_Click
        FREmail.cmdExit_Click
    End If
    Set clsSMTP = Nothing
End Sub

Public Function VerificaCedula(ByRef Cedula As String) As Boolean
    VerificaCedula = True
    Tipo = "C"
    If Len(Trim(Cedula)) <> 10 And Len(Trim(Cedula)) <> 13 Then
        VerificaCedula = False
    End If

    If Val(Mid(Cedula, 1, 2)) > 25 Then
        VerificaCedula = False
    End If

    If Val(Mid(Cedula, 3, 1)) > 5 Then
        VerificaCedula = True
        Tipo = "R"
    End If

    If VerificaCedula = False Then
        VerificaCedula = False
    Else
        Dim Total As Integer
        Dim Cifra As Integer
        Total = 0
        If Mid(Cedula, 3, 1) <> 6 Then
            If Tipo = "C" Then
                For a = 1 To 9
                    If (a Mod 2) = 0 Then
                        Cifra = Val(Mid(Cedula, a, 1))
                    Else
                    Cifra = Val(Mid(Cedula, a, 1)) * 2
                        If Cifra > 9 Then
                            Cifra = Cifra - 9
                        End If
                    End If
                    Total = Total + Cifra
                Next
            
                Cifra = Total Mod 10
            
                If Cifra > 0 Then
                    Cifra = 10 - Cifra
                End If
            ElseIf Tipo = "R" Then
                For a = 1 To 9
                    
                    
                    '4     3     2    7   6   5   4   3   2
                    Total = Total + Val(Mid(Cedula, a, 1)) * IIf((11 - a) <= 7, (11 - a), (5 - a))
                Next
            
                Cifra = Total Mod 11
            
                If Cifra > 0 Then
                    Cifra = 11 - Cifra
                End If
            
            End If
        
            If Cifra = Val(Mid(Cedula, 10, 1)) Then
                VerificaCedula = True
    '            MsgBox "Numero de cedula SI pasa la validacin, verifique por favor", vbInformation
            Else
                MsgBox "Numero de cedula NO pasa la validacin, verifique por favor", vbInformation
                VerificaCedula = False
            End If
        Else
        
            For a = 1 To 8
                '4     3     2    7   6   5   4   3   2
                Total = Total + Val(Mid(Cedula, a, 1)) * IIf((10 - a) <= 7, (10 - a), (4 - a))
            Next
        
            Cifra = Total Mod 11
        
            If Cifra > 0 Then
                Cifra = 11 - Cifra
            End If
            If Cifra = Val(Mid(Cedula, 9, 1)) Then
                VerificaCedula = True
    '            MsgBox "Numero de cedula SI pasa la validacin, verifique por favor", vbInformation
            Else
                MsgBox "Numero de cedula NO pasa la validacin, verifique por favor", vbInformation
                VerificaCedula = False
            End If
        End If
    End If
    If VerificaCedula = False Then
        If MsgBox("Cedula incorrecta." & vbNewLine & "Es Pasaporte?", vbQuestion + vbYesNo + vbDefaultButton2, "Verificación CI/RUC") = vbYes Then
            Cedula = "P" & Cedula
            VerificaCedula = True
        End If
    End If
End Function

Public Function SumaDiasHabiles(strFecha As String, sumaDias As Long) As String
    Dim strFechaSumada As String
    strFechaSumada = strFecha
    sumaDias = sumaDias - 1
    For i = 1 To sumaDias
        strFechaSumada = DateAdd("d", 1, strFechaSumada)
        If Weekday(strFechaSumada, vbMonday) = 6 Or Weekday(strFechaSumada, vbMonday) = 7 Then
            i = i - 1
        End If
    Next i
    SumaDiasHabiles = strFechaSumada
End Function


Public Sub DocElectronico(strCodDoc As String, codDoc As String)
    Dim clsCon_Aux As New clsConsulta
    Dim clscon_Ing As New clsConsulta
    Dim strSql As String
    Dim Valor As String
    If GeneraDocElec = 1 Then
        clsCon_Aux.Inicializar AdoConn, AdoConnMaster
        clscon_Ing.Inicializar AdoConn, AdoConnMaster
        If strCodDoc = "01" Then
            strSql = " SELECT campo " & _
                     " FROM egreso INNER JOIN doc_electronico_campoadicional " & _
                     " ON egreso.emp_codigo=doc_electronico_campoadicional.emp_codigo " & _
                     " AND egreso.per_codigo=doc_electronico_campoadicional.per_codigo " & _
                     " AND doc_electronico_campoadicional.doc_ele_coddoc='01' " & _
                     " WHERE egreso.emp_codigo='" & strEmpresa & "' " & _
                     " AND egreso.tip_egr_codigo='FAC' " & _
                     " AND egreso.egr_codigo='" & codDoc & "' "
            clsCon_Aux.Ejecutar strSql
            If clsCon_Aux.adorec_Def.RecordCount > 0 Then
                While Not clsCon_Aux.adorec_Def.EOF
                    While Valor = ""
                        Valor = InputBox("Ingrese el valor para el campo" & vbNewLine & _
                                clsCon_Aux.adorec_Def("campo") & vbNewLine & _
                                "Campo requerido por el cliente para el archivo XML", "CampoAdicional")
                        If Valor = "" Then
                            If MsgBox("No desea añadir el campo al archivo XML?", vbYesNo + vbDefaultButton2 + vbQuestion, "CampoAdicional") = vbYes Then
                                Valor = "VACIO"
                            End If
                        End If
                    Wend
                    If Valor <> "VACIO" Then
                        strSql = " INSERT INTO doc_electronico_campoadicional_documento " & _
                                 " (emp_codigo, doc_ele_coddoc, doc_ele_codigo, " & _
                                 " campo, valor, " & _
                                 " doc_ele_cam_doc_usumod, doc_ele_cam_doc_fechamod)" & _
                                 " VALUES('" & strEmpresa & "','" & strCodDoc & "','" & codDoc & "'," & _
                                 " '" & clsCon_Aux.adorec_Def("campo") & "','" & Valor & "'," & _
                                 " '" & strUsuario & "',CURRENT_TIMESTAMP)"
                        clscon_Ing.Ejecutar strSql, "M"
                    End If
                    clsCon_Aux.adorec_Def.MoveNext
                Wend
            End If
            strSql = " INSERT INTO doc_electronico (emp_codigo, doc_ele_coddoc, " & _
                     " doc_ele_codigo,doc_ele_estado,doc_ele_fechaemision) " & _
                     " SELECT egreso.emp_codigo,'" & strCodDoc & "',egr_codigo,'0',egr_fechamod " & _
                     " FROM egreso " & _
                     " WHERE egr_codigo ='" & codDoc & "' " & _
                     " AND tip_egr_codigo='FAC'" & _
                     " AND emp_codigo='" & strEmpresa & "'" & _
                     " AND egr_anulado=0 "
            clscon_Ing.Ejecutar strSql, "M"
        ElseIf strCodDoc = "04" Then
            strSql = " SELECT campo " & _
                     " FROM ingreso INNER JOIN doc_electronico_campoadicional " & _
                     " ON ingreso.emp_codigo=doc_electronico_campoadicional.emp_codigo " & _
                     " AND ingreso.per_codigo=doc_electronico_campoadicional.per_codigo " & _
                     " AND doc_electronico_campoadicional.doc_ele_coddoc='04' " & _
                     " WHERE ingreso.emp_codigo='" & strEmpresa & "' " & _
                     " AND ingreso.tip_ing_codigo='DCL' " & _
                     " AND ingreso.ing_codigo='" & codDoc & "' "
            clsCon_Aux.Ejecutar strSql
            If clsCon_Aux.adorec_Def.RecordCount > 0 Then
                While Not clsCon_Aux.adorec_Def.EditMode
                    While Valor = ""
                        Valor = InputBox("Ingrese el valor para el campo" & vbNewLine & _
                                clsCon_Aux.adorec_Def("campo") & vbNewLine & _
                                "Campo requerido por el cliente para el archivo XML", "CampoAdicional")
                        If Valor = "" Then
                            If MsgBox("No desea añadir el campo al archivo XML?", vbYesNo + vbDefaultButton2 + vbQuestion, "CampoAdicional") = vbYes Then
                                Valor = "VACIO"
                            End If
                        End If
                    Wend
                    If Valor <> "VACIO" Then
                        strSql = " INSERT INTO doc_electronico_campoadicional_documento " & _
                                 " (emp_codigo, doc_ele_coddoc, doc_ele_codigo, " & _
                                 " campo, valor, " & _
                                 " doc_ele_cam_doc_usumod, doc_ele_cam_doc_fechamod)" & _
                                 " VALUES('" & strEmpresa & "','" & strCodDoc & "','" & codDoc & "'," & _
                                 " '" & clsCon_Aux.adorec_Def("campo") & "','" & Valor & "'," & _
                                 " '" & strUsuario & "',CURRENT_TIMESTAMP)"
                        clscon_Ing.Ejecutar strSql, "M"
                    End If
                    clsCon_Aux.adorec_Def.MoveNext
                Wend
            End If
            strSql = " INSERT INTO doc_electronico (emp_codigo, doc_ele_coddoc, " & _
                     " doc_ele_codigo,doc_ele_estado,doc_ele_fechaemision) " & _
                     " SELECT ingreso.emp_codigo,'" & strCodDoc & "',ing_codigo,'0',ing_fechamod " & _
                     " FROM ingreso " & _
                     " WHERE ing_codigo='" & codDoc & "' " & _
                     " AND tip_ing_codigo='DCL'" & _
                     " AND emp_codigo='" & strEmpresa & "'" & _
                     " AND ing_anulado=0 "
            clscon_Ing.Ejecutar strSql, "M"
        ElseIf strCodDoc = "06" Then
            strSql = " INSERT INTO doc_electronico (emp_codigo, doc_ele_coddoc, " & _
                     " doc_ele_codigo,doc_ele_estado,doc_ele_fechaemision) " & _
                     " SELECT egreso.emp_codigo,'" & strCodDoc & "',egr_gui_codigo,'0',egr_fechamod " & _
                     " FROM egreso_guia INNER JOIN egreso ON egreso_guia.emp_codigo=egreso.emp_codigo" & _
                     " AND egreso_guia.tip_egr_codigo=egreso.tip_egr_codigo" & _
                     " AND egreso_guia.egr_codigo=egreso.egr_codigo" & _
                     " WHERE egr_gui_codigo ='" & codDoc & "' " & _
                     " AND egreso_guia.emp_codigo='" & strEmpresa & "'" & _
                     " AND egr_anulado=0 "
            clscon_Ing.Ejecutar strSql, "M"
    
        ElseIf strCodDoc = "07" Then
            strSql = " SELECT campo " & _
                     " FROM cuenta_p_c INNER JOIN doc_electronico_campoadicional " & _
                     " ON cuenta_p_c.emp_codigo=doc_electronico_campoadicional.emp_codigo " & _
                     " AND cuenta_p_c.per_codigo=doc_electronico_campoadicional.per_codigo " & _
                     " AND doc_electronico_campoadicional.doc_ele_coddoc='07' " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='P' " & _
                     " AND cuenta_p_c.cue_p_c_codigo='" & codDoc & "' "
            clsCon_Aux.Ejecutar strSql
            If clsCon_Aux.adorec_Def.RecordCount > 0 Then
                While Not clsCon_Aux.adorec_Def.EOF
                    While Valor = ""
                        Valor = InputBox("Ingrese el valor para el campo" & vbNewLine & _
                                clsCon_Aux.adorec_Def("campo") & vbNewLine & _
                                "Campo requerido por el cliente para el archivo XML", "CampoAdicional")
                        If Valor = "" Then
                            If MsgBox("No desea añadir el campo al archivo XML?", vbYesNo + vbDefaultButton2 + vbQuestion, "CampoAdicional") = vbYes Then
                                Valor = "VACIO"
                            End If
                        End If
                    Wend
                    If Valor <> "VACIO" Then
                        strSql = " INSERT INTO doc_electronico_campoadicional_documento " & _
                                 " (emp_codigo, doc_ele_coddoc, doc_ele_codigo, " & _
                                 " campo, valor, " & _
                                 " doc_ele_cam_doc_usumod, doc_ele_cam_doc_fechamod)" & _
                                 " VALUES('" & strEmpresa & "','" & strCodDoc & "','" & codDoc & "'," & _
                                 " '" & clsCon_Aux.adorec_Def("campo") & "','" & Valor & "'," & _
                                 " '" & strUsuario & "',CURRENT_TIMESTAMP)"
                        clscon_Ing.Ejecutar strSql, "M"
                    End If
                    clsCon_Aux.adorec_Def.MoveNext
                Wend
            End If
            strSql = " INSERT INTO doc_electronico (emp_codigo, doc_ele_coddoc," & _
                     " doc_ele_codigo,doc_ele_estado,doc_ele_fechaemision) " & _
                     " SELECT comprobante_retencion.emp_codigo,'" & strCodDoc & "',cue_p_c_codigo,'0',com_ret_fechamod" & _
                     " FROM comprobante_retencion " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='P'" & _
                     " AND cue_p_c_codigo='" & codDoc & "'" & _
                     " AND com_ret_total!=0 "
            clscon_Ing.Ejecutar strSql, "M"
        End If
    End If
End Sub

Public Function Buscar_Carpeta(ByVal hWndOwner As Long, Optional Titulo As String, Optional Path_Inicial As Variant) As String

On Local Error GoTo errFunction
    
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            hWndOwner, _
                            Titulo, _
                            0, _
                            17)
    
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
    
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path

Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString

End Function

Public Function DisponibleParaLaVentaBodega(producto As String, bodega As String) As Long
    Dim clsExis As New clsConsulta
    Dim SQL As String
    clsExis.Inicializar AdoConn, AdoConnMaster
    SQL = " SELECT producto.prd_codigo,sum(exi) as EXISTENCIA " & _
          " From Producto Inner Join" & _
          " (SELECT p.emp_codigo, p.prd_codigo, Sum(e.exi_cantidad) AS exi" & _
          " FROM producto p inner join existencia e" & _
          " on p.emp_codigo=e.emp_codigo" & _
          " and p.prd_codigo=e.prd_codigo" & _
          " and e.dep_codigo LIKE '" & bodega & "'" & _
          " WHERE p.emp_codigo='RYB'" & _
          " AND p.prd_codigo='" & producto & "'" & _
          " GROUP BY p.emp_codigo, p.prd_codigo"
    SQL = SQL & " Union" & _
          " SELECT p.emp_codigo,p.prd_codigo, -1*Sum(det_pedido.det_ped_cant_pedida) as exi" & _
          " FROM pedido INNER JOIN det_pedido" & _
          " ON pedido.emp_codigo=det_pedido.emp_codigo" & _
          " AND pedido.ped_codigo=det_pedido.ped_codigo" & _
          " and det_pedido.dep_codigo LIKE '" & bodega & "'" & _
          " INNER JOIN producto p" & _
          " on det_pedido.emp_codigo=p.emp_codigo" & _
          " and det_pedido.prd_codigo=p.prd_codigo" & _
          " AND p.prd_codigo='" & producto & "'" & _
          " WHERE pedido.emp_codigo='RYB'" & _
          " AND pedido.ped_estado in (0,1)" & _
          " GROUP BY p.emp_codigo, p.prd_codigo"
    SQL = SQL & " ) e" & _
          " on producto.emp_codigo=e.emp_codigo" & _
          " and producto.prd_codigo=e.prd_codigo" & _
          " WHERE producto.emp_codigo='RYB'" & _
          " AND producto.prd_codigo='" & producto & "'" & _
          " GROUP BY producto.prd_codigo"
    clsExis.Ejecutar SQL
    If clsExis.adorec_Def.RecordCount > 0 Then
        DisponibleParaLaVentaBodega = (clsExis.adorec_Def("EXISTENCIA"))
    Else
        DisponibleParaLaVentaBodega = 0
    End If
End Function

Public Sub InicializarContenedorRecurrente()
    booUnContenedor = False
    strContenedorRecurrente = ""
End Sub

Public Function Hoy() As String
    On Error GoTo errhandler
    Set clsHoy = New clsConsulta
    clsHoy.Inicializar AdoConn, AdoConnMaster
    strSql = "SELECT CURRENT_TIMESTAMP as hoy"
    clsHoy.Ejecutar strSql
    Hoy = Left(clsHoy.adorec_Def("hoy"), 10)
    Set clsHoy = Nothing
    Exit Function
errhandler:
    Hoy = Left(Now, 10)
    Set clsHoy = Nothing
End Function

Public Function Ahora() As String
    Set clsAhora = New clsConsulta
    On Error GoTo errhandler
    clsAhora.Inicializar AdoConn, AdoConnMaster
    strSql = "SELECT CURRENT_TIMESTAMP as hoy"
    clsAhora.Ejecutar strSql
    Ahora = clsAhora.adorec_Def("hoy")
    Set clsAhora = Nothing
    Exit Function
errhandler:
    Ahora = Now
    Set clsAhora = Nothing
End Function

Public Sub RegistraError(Pantalla As String, Descripcion As String, CadenaSQL As String)
    Dim SQL As String
    Dim clsAux As New clsConsulta
    clsAux.Inicializar AdoConn, AdoConnMaster
    SQL = " INSERT INTO error (err_pantalla,err_descripcion,err_sql,err_fechamod,err_usumod) " & _
          " VALUES('" & Pantalla & "','" & Replace(Descripcion, "'", "`") & "','" & Replace(CadenaSQL, "'", "`") & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
    clsAux.Ejecutar SQL, "M"
    Set clsAux = Nothing
End Sub



Public Function RevisivaCodigoIVAFactura(strFactura As String) As Integer
    Dim clscon As New clsConsulta
    clscon.Inicializar AdoConn, AdoConnMaster
    RevisivaCodigoIVAFactura = 0
    
    strSql = " SELECT cod_iva_codigo FROM egreso " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND tip_egr_codigo='FAC' " & _
             " AND egr_codigo='" & strFactura & "'"
    clscon.Ejecutar strSql
    If clscon.adorec_Def.RecordCount > 0 Then
        RevisivaCodigoIVAFactura = clscon.adorec_Def("cod_iva_codigo")
    End If
    Set clscon = Nothing
End Function

Public Function EgresoRet(ped As String) As Boolean
    Dim clsAuxver As New clsConsulta
    clsAuxver.Inicializar AdoConn, AdoConnMaster
    EgresoRet = False
    If ped <> "" Then
        strSql = " SELECT count(*) as pasa " & _
                 " FROM pedido inner join no ON pedido.emp_codigo=no.emp_codigo " & _
                 " AND pedido.ped_tip_egr_codigo=no.tip_egr_codigo " & _
                 " AND pedido.ped_egr_codigo=no.egr_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND pedido.ped_codigo='" & ped & "' "
        clsAuxver.Ejecutar strSql
        If clsAuxver.adorec_Def.RecordCount > 0 Then
            If clsAuxver.adorec_Def("pasa") > 0 Then EgresoRet = True
        End If
    End If

End Function

Public Sub LiberarYBajarPedidos(NoBajarPedidos As Boolean, strTipoPedidos As String)
    Dim clsConPed As New clsConsulta
    Dim clsConLiberar As New clsConsulta
    Dim clsConNC As New clsConsulta
    clsConPed.Inicializar AdoConn, AdoConnMaster
    clsConLiberar.Inicializar AdoConn, AdoConnMaster
    clsConNC.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT pedido.emp_codigo,pedido.ped_codigo,pedido.per_codigo,  " & _
             " ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
             " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
             " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
             " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) " & _
             "-COALESCE(doc_pag_valor,0.00),2) as d " & _
             " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
             " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo in (" & strTipoPedidos & ")" & _
             " INNER JOIN det_pedido ON pedido.emp_codigo=det_pedido.emp_codigo AND pedido.ped_codigo=det_pedido.ped_codigo  " & _
             " INNER JOIN producto ON det_pedido.emp_codigo=producto.emp_codigo AND det_pedido.prd_codigo=producto.prd_codigo AND prd_incentivo=0 " & _
             " INNER JOIN parametro ON pedido.emp_codigo=parametro.emp_codigo AND parametro.par_codigo='IVAV' " & _
             " LEFT JOIN producto_promo ON det_pedido.prd_codigo=producto_promo.prd_codigo AND det_pedido.emp_codigo=producto_promo.emp_codigo " & _
             " AND LEFT(pedido.ped_fechamod,10) BETWEEN producto_promo.prd_pro_fechaini AND producto_promo.prd_pro_fechafin AND producto_promo.tip_ped_codigo=persona.tip_ped_codigo " & _
             " LEFT JOIN producto_promo2 ON det_pedido.prd_codigo=producto_promo2.prd_codigo AND det_pedido.emp_codigo=producto_promo2.emp_codigo " & _
             " AND pedido.ped_codigo=producto_promo2.ped_codigo " & _
             " LEFT JOIN (SELECT emp_codigo,ped_codigo,per_codigo,SUM(doc_pag_ped_valor) as doc_pag_valor" & _
             " FROM doc_pago_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "' AND doc_pag_ped_estado='GIRADO'" & _
             " GROUP BY emp_codigo,ped_codigo,per_codigo) pag " & _
             " ON pedido.emp_codigo=pag.emp_codigo AND pedido.ped_codigo=pag.ped_codigo " & _
             " AND pedido.per_codigo=pag.per_codigo "
    strSql = strSql & " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
             " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
             " GROUP BY pedido.emp_codigo,pedido.ped_codigo,pedido.per_codigo,par_numero,doc_pag_valor,pedido.ped_dctoadicional " & _
             " HAVING ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2)" & _
             " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2)) " & _
             " + ROUND(SUM(ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2)))),2) " & _
             " - ROUND((((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio) - (IIF(det_ped_dcto=0 OR COALESCE(pro_pre_mon_dct_dcto,0.00)!=0,ROUND((det_ped_cant_entregada+det_ped_cant_programada)*det_ped_precio*IIF(IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00))>COALESCE(per_dcto,0),IIF(COALESCE(prd_pro_porcentaje,0.00)>=COALESCE(pro_pre_mon_dct_dcto,0.00),COALESCE(prd_pro_porcentaje,0.00),COALESCE(pro_pre_mon_dct_dcto,0.00)),COALESCE(per_dcto,0))/100.00,2),ROUND(det_ped_dcto/det_ped_cant_pedida*(det_ped_cant_entregada+det_ped_cant_programada),2))))*(pedido.ped_dctoadicional/100.00),2))* (par_numero)/100.00,2) " & _
             "-COALESCE(doc_pag_valor,0.00),2)<0.02 "
    clsConPed.Ejecutar strSql
    While Not clsConPed.adorec_Def.EOF
        strSql = " UPDATE pedido " & _
                 " SET ped_estado=1 " & _
                 " WHERE emp_codigo='" & clsConPed.adorec_Def("emp_codigo") & "'" & _
                 " AND ped_codigo='" & clsConPed.adorec_Def("ped_codigo") & "'" & _
                 " AND per_codigo='" & clsConPed.adorec_Def("per_codigo") & "'"
        clsConLiberar.Ejecutar strSql, "M"
        clsConPed.adorec_Def.MoveNext
    Wend
    
    If NoBajarPedidos = False Then
        strSql = " SELECT pedido.emp_codigo,pedido.ped_codigo,pedido.per_codigo " & _
                 " FROM pedido INNER JOIN persona ON pedido.emp_codigo=persona.emp_codigo" & _
                 " AND pedido.per_codigo=persona.per_codigo AND persona.tip_ped_codigo in (" & strTipoPedidos & ")" & _
                 " WHERE pedido.emp_codigo = '" & strEmpresa & "' " & _
                 " AND persona.for_pag_codigo in ('EFE','CONT') AND pedido.ped_estado in (0)" & _
                 " AND ped_fechamod <= IIF(DATEPART(dw,CURRENT_TIMESTAMP) in (2,3,4), DATEADD(d,-5,CURRENT_TIMESTAMP), DATEADD(d,-3,CURRENT_TIMESTAMP))"
'        strSQL = "select pedido.emp_codigo,pedido.ped_codigo,pedido.per_codigo from pedido inner join (select distinct ped_codigo from doc_pago_pedido where doc_pag_ped_codigo like 'DCL%' and doc_pag_ped_estado='GIRADO') pp on pedido.ped_codigo=pp.ped_codigo where pedido.ped_estado=3"
        clsConPed.Ejecutar strSql
        While Not clsConPed.adorec_Def.EOF
            strSql = " UPDATE pedido " & _
                     " SET ped_estado=9, " & _
                     " ped_fechamod=CURRENT_TIMESTAMP, " & _
                     " ped_usumod='" & strUsuario & "' " & _
                     " WHERE emp_codigo='" & clsConPed.adorec_Def("emp_codigo") & "'" & _
                     " AND ped_codigo='" & clsConPed.adorec_Def("ped_codigo") & "'" & _
                     " AND per_codigo='" & clsConPed.adorec_Def("per_codigo") & "'"
            clsConLiberar.Ejecutar strSql, "M"
            
            LiberarIncentivo clsConPed.adorec_Def("ped_codigo"), clsConPed.adorec_Def("per_codigo")
            
            clsConPed.adorec_Def.MoveNext
        Wend
    End If
End Sub

Public Sub LiberarIncentivo(strPedido As String, Optional strCodCliente As String = "")
    Dim clsConLiberar As New clsConsulta
    Dim clsConNC As New clsConsulta
    clsConLiberar.Inicializar AdoConn, AdoConnMaster
    clsConNC.Inicializar AdoConn, AdoConnMaster
    
    If strCodCliente = "" Then
        strSql = " SELECT per_codigo FROM pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' AND ped_codigo=" & strPedido
        clsConLiberar.Ejecutar (strSql), "M"
        strCodCliente = clsConLiberar.adorec_Def("per_codigo")
    End If
    strSql = " UPDATE incentivo_local " & _
             " SET inc_loc_estado=0, " & _
             " ped_codigo=null," & _
             " inc_loc_fechamod=CURRENT_TIMESTAMP," & _
             " inc_loc_usumod='" & strUsuario & "'" & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND ped_codigo='" & strPedido & "'"
    clsConLiberar.Ejecutar strSql, "M"
    strSql = " SELECT emp_codigo,doc_pag_ped_codigo,doc_pag_ped_numero,per_codigo,doc_pag_ped_valor " & _
             " FROM doc_pago_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND ped_codigo='" & strPedido & "'" & _
             " AND per_codigo='" & strCodCliente & "'" & _
             " AND doc_pag_ped_codigo like 'DCL%'" & _
             " AND doc_pag_ped_tipo='N' AND tip_doc_pag_codigo='DCL' AND doc_pag_ped_estado='GIRADO'"
    clsConNC.Ejecutar strSql
    While Not clsConNC.adorec_Def.EOF
        strSql = " UPDATE ingreso " & _
                 " SET ing_saldo=ing_saldo-'" & clsConNC.adorec_Def("doc_pag_ped_valor") & "' " & _
                 " WHERE emp_codigo='" & clsConNC.adorec_Def("emp_codigo") & "'" & _
                 " AND tip_ing_codigo='DCL'" & _
                 " AND ing_codigo='" & clsConNC.adorec_Def("doc_pag_ped_numero") & "'" & _
                 " AND per_codigo='" & clsConNC.adorec_Def("per_codigo") & "'"
        clsConLiberar.Ejecutar strSql, "M"
        strSql = " UPDATE doc_pago_pedido " & _
                 " SET " & _
                 " doc_pag_ped_estado='ANULADO'," & _
                 " doc_pag_ped_fechamod=CURRENT_TIMESTAMP," & _
                 " doc_pag_ped_usumod='" & strUsuario & "'" & _
                 " WHERE emp_codigo='" & clsConNC.adorec_Def("emp_codigo") & "'" & _
                 " AND tip_doc_pag_codigo='DCL'" & _
                 " AND doc_pag_ped_numero='" & clsConNC.adorec_Def("doc_pag_ped_numero") & "'" & _
                 " AND per_codigo='" & clsConNC.adorec_Def("per_codigo") & "'" & _
                 " AND doc_pag_ped_codigo='" & clsConNC.adorec_Def("doc_pag_ped_codigo") & "'"
        clsConLiberar.Ejecutar strSql, "M"
        clsConNC.adorec_Def.MoveNext
    Wend

End Sub

Public Sub PagarFacturaDePedidoPagado(strPedido As String, strFactura As String, strPedReprogramado As String)
    Dim clsConDoc As New clsConsulta
    Dim clsConFac As New clsConsulta
    Dim clsConPed As New clsConsulta
    Dim clsConAUX As New clsConsulta
    Dim clsAsientoE As New clsContable
    Dim ValAplica As Double
    Dim ValAplicaRepro As Double
    Dim ValPago As Double
    Dim Pag As Double
    Dim maxpag As Long
    Dim CtaCXC As String
    Dim Desc As String
    Dim dblAnticipo As Double
    Dim i As Long
    Dim MaxPedidos As Long
    clsConDoc.Inicializar AdoConn, AdoConnMaster
    clsConFac.Inicializar AdoConn, AdoConnMaster
    clsConPed.Inicializar AdoConn, AdoConnMaster
    clsConAUX.Inicializar AdoConn, AdoConnMaster
    clsAsientoE.Inicializar AdoConn, AdoConnMaster
    If Len(strFactura) > 0 Then
        strFactura = Left(strFactura, Len(strFactura) - 1)
    
    '***************COBROS
        strSql = " SELECT emp_codigo,doc_pag_ped_tipo,doc_pag_ped_codigo,doc_pag_codigo," & _
                 " tip_doc_pag_codigo,ban_codigo,doc_pag_ped_numero,doc_pag_ped_fecha_recepcion," & _
                 " doc_pag_ped_fecha_doc,doc_pag_ped_fecha_efec,per_codigo,doc_pag_ped_valor," & _
                 " doc_pag_ped_observacion,doc_pag_ped_estado,doc_pag_ped_pendiente,doc_pag_ped_anticipo," & _
                 " doc_pag_ped_saldo,ped_codigo,per_codigo_ch,doc_pag_ped_fechamod,doc_pag_ped_usumod" & _
                 " FROM doc_pago_pedido " & _
                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo='" & strPedido & "' " & _
                 " AND doc_pag_ped_codigo like 'CSP%' " & _
                 " AND doc_pag_ped_tipo!='N' " & _
                 " AND doc_pag_ped_estado NOT IN ('ANULADO','APLICADO') " & _
                 " AND doc_pag_ped_valor-doc_pag_ped_saldo>0 " & _
                 " ORDER BY doc_pag_ped_fecha_recepcion,doc_pag_ped_codigo"
        clsConDoc.Ejecutar strSql, "M"
        While Not clsConDoc.adorec_Def.EOF
                '*************************************************
            ValAplica = clsConDoc.adorec_Def("doc_pag_ped_valor")
            'Calcula el máximo codigo de pago para la cuenta
            strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_fechaemision,(cue_p_c_valor-COALESCE(sum(pag_monto),0)) as cue_p_c_valor " & _
                     " FROM cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_egr_codigo IN ('" & Replace(strFactura, ",", "','") & "') " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C' " & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_valor,cuenta_p_c.cue_p_c_fechaemision " & _
                     " ORDER BY cuenta_p_c.cue_p_c_codigo"
                     
            clsConPed.Ejecutar strSql, "M"
            
            dblAnticipo = FormatoD2(0)
            While (Not clsConPed.adorec_Def.EOF) And ValAplica > 0
'*******************ANTICIPO
                dblAnticipo = FormatoD2(0)
                If FormatoD2(ValAplica) > FormatoD2(clsConPed.adorec_Def("cue_p_c_valor")) Then
                    dblAnticipo = FormatoD2(FormatoD2(ValAplica) - FormatoD2(clsConPed.adorec_Def("cue_p_c_valor")))
                    
                    If strPedReprogramado <> "" Then
                        ValAplicaRepro = ValAplica
                        ReDim PedReprogramado(UBound(Split(strPedReprogramado, ","))) As String
                        PedReprogramado = Split(strPedReprogramado, ",")
                        MaxPedidos = UBound(Split(strPedReprogramado, ","))
                        
                        For i = 0 To MaxPedidos
                            strSql = " SELECT ped_subtotal " & _
                                     " FROM pedido " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                     " AND ped_codigo='" & PedReprogramado(i) & "'"
                            clsConAUX.Ejecutar strSql, "M"
                            ValAplicaRepro = ValAplicaRepro - FormatoD2(clsConAUX.adorec_Def("ped_subtotal"))
                            strSql = " INSERT INTO doc_pago_pedido (emp_codigo,doc_pag_ped_tipo,doc_pag_ped_codigo," & _
                                     " tip_doc_pag_codigo,ban_codigo,doc_pag_ped_numero," & _
                                     " doc_pag_ped_fecha_recepcion,doc_pag_ped_fecha_doc," & _
                                     " doc_pag_ped_fecha_efec,per_codigo,doc_pag_codigo," & _
                                     " doc_pag_ped_valor,doc_pag_ped_observacion," & _
                                     " doc_pag_ped_estado,doc_pag_ped_pendiente,doc_pag_ped_anticipo," & _
                                     " doc_pag_ped_saldo,ped_codigo,per_codigo_ch," & _
                                     " doc_pag_ped_fechamod,doc_pag_ped_usumod)" & _
                                     " VALUES('" & strEmpresa & "','R','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-R" & i & "'," & _
                                     " '" & clsConDoc.adorec_Def("tip_doc_pag_codigo") & "', '" & clsConDoc.adorec_Def("ban_codigo") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', " & _
                                     " '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "','" & clsConDoc.adorec_Def("doc_pag_ped_fecha_doc") & "'," & _
                                     " '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', '" & clsConDoc.adorec_Def("per_codigo") & "','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-A'," & _
                                     " '" & FormatoD2(clsConAUX.adorec_Def("ped_subtotal")) & "', 'ANTICIPO REPROGRAMADO - " & clsConDoc.adorec_Def("doc_pag_ped_observacion") & "'," & _
                                     " '" & clsConDoc.adorec_Def("doc_pag_ped_estado") & "', 0, 1, " & _
                                     " 0,'" & PedReprogramado(i) & "','" & clsConDoc.adorec_Def("per_codigo_ch") & "'," & _
                                     " CURRENT_TIMESTAMP,'" & strUsuario & "')"
                            clsConAUX.Ejecutar strSql, "M"
                        
                        Next i
                        
                    End If
                    
                    strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, " & _
                             " tip_doc_pag_codigo, ban_codigo, " & _
                             " doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_efec, " & _
                             " doc_pag_fecha_doc , per_codigo, " & _
                             " doc_pag_valor, doc_pag_observacion, " & _
                             " doc_pag_estado,doc_pag_pendiente, " & _
                             " doc_pag_anticipo, doc_pag_fechamod, doc_pag_usumod,per_codigo_ch)" & _
                             " VALUES ('" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-A', '" & strEmpresa & "', " & _
                             " '" & clsConDoc.adorec_Def("tip_doc_pag_codigo") & "', '" & clsConDoc.adorec_Def("ban_codigo") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_doc") & "', '" & clsConDoc.adorec_Def("per_codigo") & "', " & _
                             " '" & FormatoD2(dblAnticipo) & "', 'ANTICIPO - " & IIf(strPedReprogramado = "", "", " PARA PEDIDOS " & strPedReprogramado & " - ") & clsConDoc.adorec_Def("doc_pag_ped_observacion") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_estado") & "', 0, " & _
                             " 1, '" & clsConDoc.adorec_Def("doc_pag_ped_fechamod") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_usumod") & "', '" & clsConDoc.adorec_Def("per_codigo_ch") & "') "
                    clsConAUX.Ejecutar strSql, "M"

                    strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-A',0, '*', '','" & FormatoD2(dblAnticipo) & "', 0 , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsConAUX.Ejecutar strSql, "M"

                    strSql = " SELECT IIF(cat_p_ctaconta_ant IS NULL OR cat_p_ctaconta_ant='',par_texto,cat_p_ctaconta_ant) as par_texto " & _
                             " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                             " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                             " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                             " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                             " AND per_codigo='" & clsConDoc.adorec_Def("per_codigo") & "' AND persona.cat_p_tipo='C' "
                    clsConAUX.Ejecutar strSql
                    CtaCXC = clsConAUX.adorec_Def("par_texto")
                    strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-A',0, '" & CtaCXC & "', '', 0,'" & FormatoD2(dblAnticipo) & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                    clsConAUX.Ejecutar strSql, "M"

'**************************************************************
                    clsAsientoE.NuevoAsiento "I", clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion"), 0, 0, FormatoD2(dblAnticipo), "ANTICIPO - " & IIf(strPedReprogramado = "", "", " PARA PEDIDOS " & strPedReprogramado & " - ") & clsConDoc.adorec_Def("doc_pag_ped_observacion") & vbNewLine & "NO: " & clsConDoc.adorec_Def("doc_pag_ped_numero") & " VALOR: " & FormatoD2(dblAnticipo)

                    strSql = " UPDATE doc_pago " & _
                             " SET asi_numasiento='" & clsAsientoE.NumAsiento & _
                             "' , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                             " WHERE doc_pag_codigo= '" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "-A' AND emp_codigo = '" & strEmpresa & "' "
                    clsConAUX.Ejecutar strSql, "M"
                    'ingreso del detalle del asiento
                    strSql = " SELECT cta_ban_ctaconta " & _
                             " FROM tipo_doc_pago INNER JOIN cta_banco " & _
                             " ON tipo_doc_pago.ban_codigo=cta_banco.ban_codigo " & _
                             " AND tipo_doc_pago.cta_ban_numero=cta_banco.cta_ban_numero " & _
                             " WHERE cta_banco.emp_codigo='" & strEmpresa & "' " & _
                             " AND tip_doc_pag_codigo='" & clsConDoc.adorec_Def("tip_doc_pag_codigo") & "' "
                    clsConAUX.Ejecutar strSql
                    'banco
                    clsAsientoE.NuevoDetAsiento clsConAUX.adorec_Def("cta_ban_ctaconta"), "", FormatoD2(dblAnticipo), FormatoD2("0")
                    'cliente
                    clsAsientoE.NuevoDetAsiento CtaCXC, "", FormatoD2("0"), FormatoD2(dblAnticipo)
'*******************FIN ANTICIPO
                End If
                If clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R" Then
                    strSql = " INSERT INTO doc_pago (doc_pag_codigo, emp_codigo, " & _
                             " tip_doc_pag_codigo, ban_codigo, " & _
                             " doc_pag_numero, doc_pag_fecha_recepcion, doc_pag_fecha_efec, " & _
                             " doc_pag_fecha_doc , per_codigo, " & _
                             " doc_pag_valor, doc_pag_observacion, " & _
                             " doc_pag_estado,doc_pag_pendiente, " & _
                             " doc_pag_anticipo, doc_pag_fechamod, doc_pag_usumod,per_codigo_ch)" & _
                             " VALUES ('" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "', '" & strEmpresa & "', " & _
                             " '" & clsConDoc.adorec_Def("tip_doc_pag_codigo") & "', '" & clsConDoc.adorec_Def("ban_codigo") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_doc") & "', '" & clsConDoc.adorec_Def("per_codigo") & "', " & _
                             " '" & FormatoD2(ValAplica) - FormatoD2(dblAnticipo) & "', '" & clsConDoc.adorec_Def("doc_pag_ped_observacion") & vbNewLine & " FAC: " & strFactura & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_estado") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_pendiente") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_anticipo") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_fechamod") & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_usumod") & "', '" & clsConDoc.adorec_Def("per_codigo_ch") & "') "
                    clsConAUX.Ejecutar strSql, "M"
                End If
                  
                strSql = " SELECT COALESCE(max(pag_codigo),0) as pag " & _
                         " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                         " AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                         " AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                         " WHERE cuenta_p_c.cue_p_c_codigo= '" & clsConPed.adorec_Def("cue_p_c_codigo") & "' " & _
                         " AND pago.emp_codigo = '" & strEmpresa & "' AND pago.cue_p_c_tipo = 'C'" & _
                         " GROUP BY pago.emp_codigo"
                clsConAUX.Ejecutar strSql
                If clsConAUX.adorec_Def.EOF Then
                    maxpag = 1
                Else
                    maxpag = clsConAUX.adorec_Def("pag") + 1
                End If
                
                If clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R" Then
                    strSql = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, " & _
                             " pag_fecha, pag_monto, " & _
                             " pag_no_doc, pag_observacion," & _
                             " doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & clsConPed.adorec_Def("cue_p_c_codigo") & "', 'C', '" & Val(maxpag) & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion") & "', '" & FormatoD2(ValAplica) - FormatoD2(dblAnticipo) & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_observacion") & vbNewLine & " FAC: " & strFactura & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "',null,'" & clsConDoc.adorec_Def("doc_pag_ped_fechamod") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_usumod") & "') "
                    
                Else
                    strSql = " UPDATE doc_pago " & _
                             " SET doc_pag_saldo=doc_pag_saldo+'" & FormatoD2(ValAplica) - FormatoD2(dblAnticipo) & "'," & _
                             " doc_pag_anticipo=0 " & _
                             " WHERE doc_pag_codigo= '" & clsConDoc.adorec_Def("doc_pag_codigo") & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsConAUX.Ejecutar strSql, "M"
                    strSql = " UPDATE doc_pago " & _
                             " SET doc_pag_anticipo=IIF(doc_pag_valor<=doc_pag_saldo,0,1), " & _
                             " doc_pag_pendiente=IIF(doc_pag_valor<=doc_pag_saldo,0,1) " & _
                             " WHERE doc_pag_codigo= '" & clsConDoc.adorec_Def("doc_pag_codigo") & "' AND emp_codigo = '" & strEmpresa & "' "
                    clsConAUX.Ejecutar strSql, "M"
                    strSql = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, " & _
                             " pag_fecha, pag_monto, " & _
                             " pag_no_doc, pag_observacion," & _
                             " doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                             " VALUES ('" & strEmpresa & "', '" & clsConPed.adorec_Def("cue_p_c_codigo") & "', 'C', '" & Val(maxpag) & "', " & _
                             " '" & clsConPed.adorec_Def("cue_p_c_fechaemision") & "', '" & FormatoD2(ValAplica) - FormatoD2(dblAnticipo) & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', 'APLICACION - " & clsConDoc.adorec_Def("doc_pag_ped_observacion") & vbNewLine & " FAC: " & strFactura & "', " & _
                             " '" & clsConDoc.adorec_Def("doc_pag_codigo") & "',null,'" & clsConDoc.adorec_Def("doc_pag_ped_fechamod") & "', '" & clsConDoc.adorec_Def("doc_pag_ped_usumod") & "') "
                End If
                clsConAUX.Ejecutar strSql, "M"
                clsConPed.adorec_Def.MoveNext
            Wend
            If clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R" Then
                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "',0, '*', '','" & FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo)) & "', 0 , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsConAUX.Ejecutar strSql, "M"
                strSql = " SELECT IIF(cat_p_ctaconta IS NULL OR cat_p_ctaconta='',par_texto,cat_p_ctaconta) as par_texto " & _
                         " FROM persona INNER JOIN categoria_p ON persona.emp_codigo=categoria_p.emp_codigo AND persona.cat_p_codigo=categoria_p.cat_p_codigo " & _
                         " AND persona.cat_p_tipo=categoria_p.cat_p_tipo " & _
                         " INNER JOIN parametro ON persona.emp_codigo=parametro.emp_codigo AND par_codigo='CXC' " & _
                         " WHERE persona.emp_codigo='" & strEmpresa & "' " & _
                         " AND per_codigo='" & clsConDoc.adorec_Def("per_codigo") & "' AND persona.cat_p_tipo='C' "
                clsConAUX.Ejecutar strSql
                CtaCXC = clsConAUX.adorec_Def("par_texto")
                strSql = " INSERT INTO det_doc_pago (emp_codigo, doc_pag_codigo, det_doc_pag_n,cta_codigo,cen_cos_codigo, det_doc_pag_debe, det_doc_pag_haber, det_doc_pag_fechamod, det_doc_pag_usumod) " & _
                         " VALUES ('" & strEmpresa & "','" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "',0, '" & CtaCXC & "', '', 0,'" & FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo)) & "' , CURRENT_TIMESTAMP, '" & strUsuario & "') "
                clsConAUX.Ejecutar strSql, "M"
            End If
                    
            '**************************************************************
            clsAsientoE.NuevoAsiento IIf(clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R", "I", "J"), clsConDoc.adorec_Def("doc_pag_ped_fecha_recepcion"), 0, 0, FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo)), IIf(clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R", "", "APLICACION - ") & clsConDoc.adorec_Def("doc_pag_ped_observacion") & vbNewLine & "FAC: " & strFactura & vbNewLine & "NO: " & clsConDoc.adorec_Def("doc_pag_ped_numero") & " VALOR: " & clsConDoc.adorec_Def("doc_pag_ped_valor")
            'Actualiza asientos en pagos
            strSql = " UPDATE pago " & _
                     " SET asi_numasiento='" & clsAsientoE.NumAsiento & _
                     "' , pag_fechamod= CURRENT_TIMESTAMP, pag_usumod='" & strUsuario & "' " & _
                     " WHERE doc_pag_codigo= '" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' AND emp_codigo = '" & strEmpresa & "' " & _
                     " AND cue_p_c_tipo='C' "
            clsConAUX.Ejecutar strSql, "M"
            If clsConDoc.adorec_Def("tip_doc_pag_codigo") <> "R" Then
                strSql = " UPDATE doc_pago " & _
                         " SET asi_numasiento='" & clsAsientoE.NumAsiento & _
                         "' , doc_pag_fechamod= CURRENT_TIMESTAMP, doc_pag_usumod='" & strUsuario & "' " & _
                         " WHERE doc_pag_codigo= '" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' AND emp_codigo = '" & strEmpresa & "' "
                clsConAUX.Ejecutar strSql, "M"
            End If
    
            strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo," & _
                     " max(doc_pago.doc_pag_fecha_doc) as fecha,cuenta_p_c.cue_p_c_valor,COALESCE(sum(p2.pag_monto),0),COALESCE(com_ret_total,0)," & _
                     " cuenta_p_c.cue_p_c_valor-COALESCE(sum(p2.pag_monto),0)-COALESCE(com_ret_total,0) as saldo " & _
                     " FROM cuenta_p_c INNER JOIN pago as p1 ON cuenta_p_c.cue_p_c_codigo=p1.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=p1.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=p1.emp_codigo " & _
                     " AND p1.doc_pag_codigo='" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' " & _
                     " INNER JOIN pago as p2 ON cuenta_p_c.cue_p_c_codigo=p2.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=p2.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=p2.emp_codigo " & _
                     " INNER JOIN doc_pago ON p2.doc_pag_codigo=doc_pago.doc_pag_codigo " & _
                     " AND p2.emp_codigo=doc_pago.emp_codigo " & _
                     " AND doc_pago.doc_pag_pendiente=0 AND doc_pago.doc_pag_estado!='ANULADO' " & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C' " & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_valor,com_ret_total "
            clsConPed.Ejecutar strSql, "M"
            While Not clsConPed.adorec_Def.EOF
                If (FormatoD2(clsConPed.adorec_Def("saldo")) <= 0) Then
                    strSql = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_fechapago='" & clsConPed.adorec_Def("fecha") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE cue_p_c_tipo= 'C' " & _
                             " AND cue_p_c_codigo= '" & clsConPed.adorec_Def("cue_p_c_codigo") & _
                             "' AND cue_p_c_egr_codigo = '" & clsConPed.adorec_Def("cue_p_c_egr_codigo") & _
                             "' AND emp_codigo = '" & strEmpresa & "' "
                    clsConAUX.Ejecutar strSql, "M"
                End If
                clsConPed.adorec_Def.MoveNext
            Wend
            
            clsAsientoE.ModificarAsiento FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo)), FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo))
            'ingreso del detalle del asiento
            If clsConDoc.adorec_Def("doc_pag_ped_tipo") = "R" Then
                strSql = " SELECT cta_ban_ctaconta " & _
                         " FROM tipo_doc_pago INNER JOIN cta_banco " & _
                         " ON tipo_doc_pago.ban_codigo=cta_banco.ban_codigo " & _
                         " AND tipo_doc_pago.cta_ban_numero=cta_banco.cta_ban_numero " & _
                         " WHERE cta_banco.emp_codigo='" & strEmpresa & "' " & _
                         " AND tip_doc_pag_codigo='SAL' "
            Else
                strSql = " SELECT cta_ban_ctaconta " & _
                         " FROM tipo_doc_pago INNER JOIN cta_banco " & _
                         " ON tipo_doc_pago.ban_codigo=cta_banco.ban_codigo " & _
                         " AND tipo_doc_pago.cta_ban_numero=cta_banco.cta_ban_numero " & _
                         " WHERE cta_banco.emp_codigo='" & strEmpresa & "' " & _
                         " AND tip_doc_pag_codigo='" & clsConDoc.adorec_Def("tip_doc_pag_codigo") & "' "
            End If
            clsConAUX.Ejecutar strSql
            'banco
            clsAsientoE.NuevoDetAsiento clsConAUX.adorec_Def("cta_ban_ctaconta"), "", FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo)), FormatoD2("0")
            'cliente
            clsAsientoE.NuevoDetAsiento CtaCXC, "", FormatoD2("0"), FormatoD2(FormatoD2(ValAplica) - FormatoD2(dblAnticipo))
                                    
    '        strSql = " SELECT COALESCE(max(not_d_c_codigo),0) as n FROM nota_d_c where emp_codigo='" & strEmpresa & "' AND tip_not_d_c='C' GROUP BY emp_codigo"
    '        clsConAUX.Ejecutar strSql, "M"
    '
    '        strSql = " INSERT INTO nota_d_c (tip_not_d_c, not_d_c_codigo, cta_ban_numero, ban_codigo, emp_codigo, tip_not_codigo, not_d_c_numero, not_d_c_fecha, not_d_c_descripcion, not_d_c_monto,asi_numasiento,not_d_c_conciliado , not_d_c_fechamod, not_d_c_usumod) " & _
    '                 " VALUES ('C','" & clsConAUX.adorec_Def("n") + 1 & "', '" & dcmbCuenta.Text & "', '" & dcmbBancoE.BoundText & "', '" & strEmpresa & "','" & dcmbTipo.BoundText & "','" & VSFG2.TextMatrix(i, 10) & "','" & fechac & "','" & Descripcion & "','" & ValorPago & "','" & clsAsientoE.NumAsiento & "',0, CURRENT_TIMESTAMP, '" & strUsuario & "')"
    '        clsConAUX.Ejecutar strSql, "M"
            Set clsAsientoE = Nothing
            
            strSql = " UPDATE doc_pago_pedido " & _
                     " SET doc_pag_ped_estado='APLICADO' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_pag_ped_codigo='" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' "
            clsConAUX.Ejecutar strSql, "M"
            clsConDoc.adorec_Def.MoveNext
        Wend
    '***************FIN COBROS
    
    '***************NOTAS DE CREDITO
        strSql = " SELECT doc_pago_pedido.emp_codigo,doc_pag_ped_tipo,doc_pag_ped_codigo,doc_pag_codigo," & _
                 " tip_doc_pag_codigo,ban_codigo,doc_pag_ped_numero,doc_pag_ped_fecha_recepcion," & _
                 " doc_pag_ped_fecha_doc,doc_pag_ped_fecha_efec,doc_pago_pedido.per_codigo,doc_pag_ped_valor," & _
                 " doc_pag_ped_observacion,doc_pag_ped_estado,doc_pag_ped_pendiente,doc_pag_ped_anticipo," & _
                 " doc_pag_ped_saldo,ped_codigo,per_codigo_ch,doc_pag_ped_fechamod,doc_pag_ped_usumod," & _
                 " ing_fecha,ing_numasiento,CONCAT(per_apellido,' ',per_nombre) as cli" & _
                 " FROM doc_pago_pedido INNER JOIN persona ON doc_pago_pedido.emp_codigo=persona.emp_codigo " & _
                 " AND doc_pago_pedido.per_codigo=persona.per_codigo " & _
                 " INNER JOIN ingreso ON doc_pago_pedido.emp_codigo=ingreso.emp_codigo " & _
                 " AND doc_pago_pedido.per_codigo=ingreso.per_codigo " & _
                 " AND doc_pago_pedido.tip_doc_pag_codigo=ingreso.tip_ing_codigo " & _
                 " AND doc_pago_pedido.doc_pag_ped_numero=ingreso.ing_codigo" & _
                 " WHERE doc_pago_pedido.emp_codigo='" & strEmpresa & "' " & _
                 " AND ped_codigo='" & strPedido & "' " & _
                 " AND doc_pag_ped_codigo like 'DCL%' " & _
                 " AND doc_pag_ped_tipo='N' " & _
                 " AND tip_doc_pag_codigo='DCL' " & _
                 " AND doc_pag_ped_estado NOT IN ('ANULADO','APLICADO') " & _
                 " AND doc_pag_ped_valor-doc_pag_ped_saldo>0 " & _
                 " ORDER BY doc_pag_ped_fecha_recepcion,doc_pag_ped_codigo"
        clsConDoc.Ejecutar strSql, "M"
        While Not clsConDoc.adorec_Def.EOF
            strSql = " SELECT cuenta_p_c.cue_p_c_codigo,(cue_p_c_valor-COALESCE(sum(pag_monto),0)) as cue_p_c_valor " & _
                     " FROM cuenta_p_c LEFT JOIN pago ON cuenta_p_c.emp_codigo=pago.emp_codigo AND cuenta_p_c.cue_p_c_codigo=pago.cue_p_c_codigo AND cuenta_p_c.cue_p_c_tipo=pago.cue_p_c_tipo " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND cue_p_c_egr_codigo IN ('" & Replace(strFactura, ",", "','") & "') " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C' GROUP BY cuenta_p_c.cue_p_c_codigo,cue_p_c_valor" & _
                     " ORDER BY cuenta_p_c.cue_p_c_codigo"
            clsConPed.Ejecutar strSql
            strSql = " SELECT COALESCE(max(pag_codigo),0) as pag,COALESCE(sum(pag_monto),0) as pago " & _
                     " FROM pago INNER JOIN cuenta_p_c ON pago.cue_p_c_codigo= cuenta_p_c.cue_p_c_codigo " & _
                     " AND pago.cue_p_c_tipo = cuenta_p_c.cue_p_c_tipo " & _
                     " AND pago.emp_codigo = cuenta_p_c.emp_codigo " & _
                     " WHERE cuenta_p_c.cue_p_c_codigo= '" & clsConPed.adorec_Def("cue_p_c_codigo") & "' " & _
                     " AND pago.emp_codigo = '" & strEmpresa & "' " & _
                     " AND pago.cue_p_c_tipo = 'C'" & _
                     " GROUP BY pago.emp_codigo"
            clsConAUX.Ejecutar strSql
            If clsConAUX.adorec_Def.EOF Then
                maxpag = 1
                Pag = 0
            Else
                maxpag = clsConAUX.adorec_Def("pag") + 1
                Pag = FormatoD2(clsConAUX.adorec_Def("pago"))
            End If
                
            If FormatoD2(clsConDoc.adorec_Def("doc_pag_ped_valor")) > FormatoD2(clsConPed.adorec_Def("cue_p_c_valor")) Then
                dblAnticipo = FormatoD2(clsConDoc.adorec_Def("doc_pag_ped_valor")) - FormatoD2(clsConPed.adorec_Def("cue_p_c_valor"))
                ValorPago = FormatoD2(clsConDoc.adorec_Def("doc_pag_ped_valor")) - dblAnticipo
                Desc = Desc & " CANCELA"
                If FormatoD2(dblAnticipo) <> 0 Then
                    strSql = " UPDATE ingreso " & _
                             " SET ing_saldo=ing_saldo-" & FormatoD2(dblAnticipo) & " " & _
                             " WHERE emp_codigo='" & strEmpresa & "' " & _
                             " AND tip_ing_codigo='DCL' " & _
                             " AND ing_codigo='" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "' "
                    clsConAUX.Ejecutar strSql, "M"
                End If
            Else
                dblAnticipo = 0
                ValorPago = FormatoD2(clsConDoc.adorec_Def("doc_pag_ped_valor"))
                Desc = Desc & " ABONA"
            End If
            
            strSql = " INSERT INTO pago(emp_codigo, cue_p_c_codigo, cue_p_c_tipo, pag_codigo, pag_fecha, pag_monto, " & _
                     " pag_no_doc, pag_observacion,doc_pag_codigo, asi_numasiento, pag_fechamod, pag_usumod) " & _
                     " VALUES ('" & strEmpresa & "', '" & clsConPed.adorec_Def("cue_p_c_codigo") & "', 'C', '" & Val(maxpag) & "', '" & clsConDoc.adorec_Def("ing_fecha") & "', '" & ValorPago & "', " & _
                     " '" & clsConDoc.adorec_Def("doc_pag_ped_numero") & "', '" & UCase(clsConDoc.adorec_Def("cli") & " - " & Desc & " - NOTA DE CRÉDITO " & clsConDoc.adorec_Def("doc_pag_ped_numero")) & "', '','" & clsConDoc.adorec_Def("ing_numasiento") & "',CURRENT_TIMESTAMP, '" & strUsuario & "') "
            clsConPed.Ejecutar strSql, "M"
            strSql = " UPDATE doc_pago_pedido " & _
                     " SET doc_pag_ped_estado='APLICADO' " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND doc_pag_ped_codigo='" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' "
            clsConPed.Ejecutar strSql, "M"
                
            strSql = " SELECT cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo," & _
                     " max(doc_pago.doc_pag_fecha_doc) as fecha,cuenta_p_c.cue_p_c_valor,COALESCE(sum(p2.pag_monto),0),COALESCE(com_ret_total,0)," & _
                     " cuenta_p_c.cue_p_c_valor-COALESCE(sum(p2.pag_monto),0)-COALESCE(com_ret_total,0) as saldo " & _
                     " FROM cuenta_p_c INNER JOIN pago as p1 ON cuenta_p_c.cue_p_c_codigo=p1.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=p1.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=p1.emp_codigo " & _
                     " AND p1.doc_pag_codigo='" & clsConDoc.adorec_Def("doc_pag_ped_codigo") & "' " & _
                     " INNER JOIN pago as p2 ON cuenta_p_c.cue_p_c_codigo=p2.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=p2.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=p2.emp_codigo " & _
                     " INNER JOIN doc_pago ON p2.doc_pag_codigo=doc_pago.doc_pag_codigo " & _
                     " AND p2.emp_codigo=doc_pago.emp_codigo " & _
                     " AND doc_pago.doc_pag_pendiente=0 AND doc_pago.doc_pag_estado!='ANULADO' " & _
                     " LEFT JOIN comprobante_retencion ON cuenta_p_c.cue_p_c_codigo=comprobante_retencion.cue_p_c_codigo " & _
                     " AND cuenta_p_c.cue_p_c_tipo=comprobante_retencion.cue_p_c_tipo " & _
                     " AND cuenta_p_c.emp_codigo=comprobante_retencion.emp_codigo " & _
                     " WHERE cuenta_p_c.emp_codigo='" & strEmpresa & "' " & _
                     " AND cuenta_p_c.cue_p_c_tipo='C' " & _
                     " GROUP BY cuenta_p_c.cue_p_c_codigo,cuenta_p_c.cue_p_c_tipo,cuenta_p_c.emp_codigo,cuenta_p_c.cue_p_c_egr_codigo,cuenta_p_c.cue_p_c_valor,com_ret_total "
            clsConPed.Ejecutar strSql, "M"
            While Not clsConPed.adorec_Def.EOF
                If (FormatoD2(clsConPed.adorec_Def("saldo")) <= 0) Then
                    strSql = " UPDATE cuenta_p_c " & _
                             " SET cue_p_c_fechapago='" & clsConPed.adorec_Def("fecha") & "', cue_p_c_pagado = 1 , cue_p_c_fechamod= CURRENT_TIMESTAMP, cue_p_c_usumod='" & strUsuario & "' " & _
                             " WHERE cue_p_c_tipo= 'C' " & _
                             " AND cue_p_c_codigo= '" & clsConPed.adorec_Def("cue_p_c_codigo") & _
                             "' AND cue_p_c_egr_codigo = '" & clsConPed.adorec_Def("cue_p_c_egr_codigo") & _
                             "' AND emp_codigo = '" & strEmpresa & "' "
                    clsConAUX.Ejecutar strSql, "M"
                End If
                clsConPed.adorec_Def.MoveNext
            Wend
                
                
            clsConDoc.adorec_Def.MoveNext
        Wend
    
    '***************FIN NOTA CREDITO
    
'***********Autorizacion PEDIDO SALIDA
        
        strSql = " SELECT egr_codigo,egr_des_anulado,egr_des_observacion," & _
                 " egr_des_fechamod,egr_des_usumod " & _
                 " FROM egreso_despacho " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND tip_egr_codigo='P' " & _
                 " AND egr_codigo='" & strPedido & "'"
        clsConDoc.Ejecutar strSql, "M"
        If clsConDoc.adorec_Def.RecordCount > 0 Then
            strSql = " SELECT egr_codigo " & _
                     " FROM egreso " & _
                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                     " AND tip_egr_codigo='FAC' " & _
                     " AND egr_codigo IN (" & strFactura & ") "
            clsConPed.Ejecutar strSql, "M"
            While Not clsConPed.adorec_Def.EOF
                strSql = " INSERT INTO egreso_despacho (emp_codigo,tip_egr_codigo,egr_codigo," & _
                         " egr_des_observacion,egr_des_anulado,egr_des_fechamod,egr_des_usumod) " & _
                         " VALUES('" & strEmpresa & "','FAC','" & clsConPed.adorec_Def("egr_codigo") & "', " & _
                         " '" & clsConDoc.adorec_Def("egr_des_observacion") & vbNewLine & "(PED:" & clsConDoc.adorec_Def("egr_codigo") & ")','0',CURRENT_TIMESTAMP,'" & strUsuario & "')"
                clsConAUX.Ejecutar strSql, "M"
                clsConPed.adorec_Def.MoveNext
            Wend
        End If
    End If
End Sub

Public Function PromoPrendaPrecio(TipoPedido As String, Pedido As Double, FechaPed As String, RecalculaObligado As Boolean) As Double
    Dim clsPromo As New clsConsulta
    Dim clsConsulta As New clsConsulta
    Dim clsEjecuta As New clsConsulta
    Dim PrendasDePromo As Long
    Dim TotalPrendasDePromo As Long
    Dim NumeroDePromo As Long
    Dim booSeModifico As Boolean
    Dim Dcto As Double
    Dim clsPedido As New clsPedidos
    clsConsulta.Inicializar AdoConn, AdoConnMaster
    clsEjecuta.Inicializar AdoConn, AdoConnMaster
    clsPromo.Inicializar AdoConn, AdoConnMaster
    clsPedido.Inicializar AdoConn, AdoConnMaster
    If RecalculaObligado = True Then
        booSeModifico = True
    Else
        booSeModifico = False
    End If
    PromoPrendaPrecio = 0
    strSql = " DELETE " & _
             " FROM det_pedido " & _
             " WHERE emp_codigo='" & strEmpresa & "'" & _
             " AND ped_codigo='" & Pedido & "'" & _
             " AND prd_codigo IN (" & _
                " SELECT DISTINCT COALESCE(promo_prenda_precio.prd_codigo,'') " & _
                " FROM promo_prenda_precio INNER JOIN det_promo_prenda_precio " & _
                " ON promo_prenda_precio.emp_codigo=det_promo_prenda_precio.emp_codigo " & _
                " AND promo_prenda_precio.tip_ped_codigo=det_promo_prenda_precio.tip_ped_codigo " & _
                " AND promo_prenda_precio.pro_pre_pre_codigo=det_promo_prenda_precio.pro_pre_pre_codigo " & _
                " WHERE promo_prenda_precio.emp_codigo='" & strEmpresa & "'" & _
                " AND promo_prenda_precio.tip_ped_codigo='" & TipoPedido & "'" & _
                " AND COALESCE(promo_prenda_precio.prd_codigo,'')!='' " & _
                " AND CURRENT_TIMESTAMP BETWEEN pro_pre_pre_fecha_desde AND pro_pre_pre_fecha_hasta " & _
             ")"
    clsPromo.Ejecutar strSql, "M"
    strSql = " SELECT pro_pre_pre_codigo " & _
             " FROM promo_prenda_precio " & _
             " WHERE promo_prenda_precio.emp_codigo='" & strEmpresa & "'" & _
             " AND promo_prenda_precio.tip_ped_codigo='" & TipoPedido & "'" & _
             " AND " & FechaPed & " BETWEEN FORMAT(pro_pre_pre_fecha_desde,'yyyyMMdd') AND FORMAT(pro_pre_pre_fecha_hasta,'yyyyMMdd')"
    clsPromo.Ejecutar strSql
    While Not clsPromo.adorec_Def.EOF
        strSql = " SELECT det_pedido.dep_codigo,det_pedido.prd_codigo,det_ped_cant_pedida,det_ped_cant_entregada,det_ped_precio,pro_pre_pre_cantidad,det_pro_pre_pre_precio,pro_pre_pre_precio,pro_pre_pre_dcto,ROUND(det_ped_dcto/det_ped_cant_pedida*det_ped_cant_entregada,2) as det_ped_dcto, " & _
                 " prd_codigo_entregar,pro_pre_pre_prd_cantidad,pro_pre_pre_prd_precio,pro_pre_pre_prd_dcto " & _
                 " FROM pedido INNER JOIN det_pedido " & _
                 " ON pedido.emp_codigo=det_pedido.emp_codigo " & _
                 " AND pedido.ped_codigo=det_pedido.ped_codigo " & _
                 " INNER JOIN ( " & _
                 " SELECT promo_prenda_precio.emp_codigo,pro_pre_pre_cantidad,det_promo_prenda_precio.prd_codigo,det_pro_pre_pre_precio,pro_pre_pre_precio,pro_pre_pre_dcto," & _
                 " COALESCE(promo_prenda_precio.prd_codigo,'') as prd_codigo_entregar,pro_pre_pre_prd_cantidad,pro_pre_pre_prd_precio,pro_pre_pre_prd_dcto" & _
                 " FROM promo_prenda_precio INNER JOIN det_promo_prenda_precio " & _
                 " ON promo_prenda_precio.emp_codigo=det_promo_prenda_precio.emp_codigo " & _
                 " AND promo_prenda_precio.tip_ped_codigo=det_promo_prenda_precio.tip_ped_codigo " & _
                 " AND promo_prenda_precio.pro_pre_pre_codigo=det_promo_prenda_precio.pro_pre_pre_codigo " & _
                 " WHERE promo_prenda_precio.emp_codigo='" & strEmpresa & "'" & _
                 " AND promo_prenda_precio.pro_pre_pre_codigo='" & clsPromo.adorec_Def("pro_pre_pre_codigo") & "'" & _
                 " AND promo_prenda_precio.tip_ped_codigo='" & TipoPedido & "'" & _
                 " AND CURRENT_TIMESTAMP BETWEEN pro_pre_pre_fecha_desde AND pro_pre_pre_fecha_hasta " & _
                 " ) promo " & _
                 " ON det_pedido.emp_codigo=promo.emp_codigo " & _
                 " AND det_pedido.prd_codigo=promo.prd_codigo " & _
                 " WHERE pedido.emp_codigo='" & strEmpresa & "'" & _
                 " AND pedido.ped_codigo='" & Pedido & "'" & _
                 " ORDER BY det_ped_precio DESC,det_ped_cant_entregada ASC"
        clsConsulta.Ejecutar (strSql), "M"
        If clsConsulta.adorec_Def.RecordCount > 0 Then
            PrendasDePromo = clsConsulta.adorec_Def("pro_pre_pre_cantidad")
            TotalPrendasDePromo = 0
            While Not clsConsulta.adorec_Def.EOF
                TotalPrendasDePromo = TotalPrendasDePromo + clsConsulta.adorec_Def("det_ped_cant_entregada")
                clsConsulta.adorec_Def.MoveNext
            Wend
            If PrendasDePromo <= TotalPrendasDePromo Then
                clsConsulta.adorec_Def.MoveFirst
                NumeroDePromo = Int(TotalPrendasDePromo / PrendasDePromo) * PrendasDePromo
                
                While (Not clsConsulta.adorec_Def.EOF) And NumeroDePromo > 0
                    Dcto = 0
                    'PRECIO ESPECIAL
                    If clsConsulta.adorec_Def("pro_pre_pre_precio") <> 0 And clsConsulta.adorec_Def("pro_pre_pre_dcto") = 0 And clsConsulta.adorec_Def("prd_codigo_entregar") = "" Then
                        If FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada")) <= NumeroDePromo Then
                            Dcto = FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada")) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                                 - FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada")) * FormatoD2(clsConsulta.adorec_Def("det_pro_pre_pre_precio"))
                            If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                                Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                            End If
                            NumeroDePromo = NumeroDePromo - FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada"))
                        Else
                            Dcto = FormatoD0(NumeroDePromo) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                                 - FormatoD0(NumeroDePromo) * FormatoD2(clsConsulta.adorec_Def("det_pro_pre_pre_precio"))
                            If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                                Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                            End If
                            NumeroDePromo = 0
                        End If
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_dcto='" & FormatoD4(Dcto) & "'" & _
                                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                                 " AND ped_codigo='" & Pedido & "'" & _
                                 " AND prd_codigo='" & clsConsulta.adorec_Def("prd_codigo") & "'" & _
                                 " AND dep_codigo='" & clsConsulta.adorec_Def("dep_codigo") & "'"
                        clsEjecuta.Ejecutar strSql, "M"
                        booSeModifico = True
                    'DCTO ESPECIAL
                    ElseIf clsConsulta.adorec_Def("pro_pre_pre_precio") = 0 And clsConsulta.adorec_Def("pro_pre_pre_dcto") <> 0 And clsConsulta.adorec_Def("prd_codigo_entregar") = "" Then
                        If FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada")) <= NumeroDePromo Then
                            Dcto = FormatoD2(FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada")) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                                 * FormatoD2(clsConsulta.adorec_Def("pro_pre_pre_dcto")) / 100)
                            If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                                Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                            End If
                            NumeroDePromo = NumeroDePromo - FormatoD0(clsConsulta.adorec_Def("det_ped_cant_entregada"))
                        Else
                            Dcto = FormatoD2(FormatoD0(NumeroDePromo) * FormatoD2(clsConsulta.adorec_Def("det_ped_precio")) _
                                 * FormatoD2(clsConsulta.adorec_Def("pro_pre_pre_dcto")) / 100)
                            If Dcto < FormatoD2(clsConsulta.adorec_Def("det_ped_dcto")) Then
                                Dcto = FormatoD2(clsConsulta.adorec_Def("det_ped_dcto"))
                            End If
                            NumeroDePromo = 0
                        End If
                        strSql = " UPDATE det_pedido " & _
                                 " SET det_ped_dcto='" & FormatoD4(Dcto) & "'" & _
                                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                                 " AND ped_codigo='" & Pedido & "'" & _
                                 " AND prd_codigo='" & clsConsulta.adorec_Def("prd_codigo") & "'" & _
                                 " AND dep_codigo='" & clsConsulta.adorec_Def("dep_codigo") & "'"
                        clsEjecuta.Ejecutar strSql, "M"
                        booSeModifico = True
                    'ENTREGA PRODUCTO
                    ElseIf clsConsulta.adorec_Def("pro_pre_pre_precio") = 0 And clsConsulta.adorec_Def("pro_pre_pre_dcto") = 0 And clsConsulta.adorec_Def("prd_codigo_entregar") <> "" Then
'                        MsgBox "AAA"
                        NumeroDePromo = Int(TotalPrendasDePromo / PrendasDePromo)
                        
                        strSql = " SELECT det_ped_cant_pedida,det_ped_precio,det_ped_dcto " & _
                                 " FROM det_pedido " & _
                                 " WHERE emp_codigo='" & strEmpresa & "' " & _
                                 " AND ped_codigo=" & Pedido & " " & _
                                 " AND prd_codigo='" & clsConsulta.adorec_Def("prd_codigo_entregar") & "' " & _
                                 " AND dep_codigo='PRI' "
                        clsEjecuta.Ejecutar (strSql), "M"
                        If clsEjecuta.adorec_Def.RecordCount = 0 Then
                            'strSQL = " INSERT INTO det_pedido (emp_codigo, ped_codigo, prd_codigo, dep_codigo, det_ped_cant_pedida, " & _
                            '         " det_ped_cant_entregada, det_ped_precio,det_ped_dcto, det_ped_fechamod, det_ped_usumod,det_ped_incentivo) " & _
                            '         " VALUES ('" & strEmpresa & "'," & num & ",'" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 1) & "'," & .TextMatrix(i, 4) & "," & .TextMatrix(i, 11) & ", " & _
                            '         " " & .TextMatrix(i, 5) & "," & .TextMatrix(i, 6) & ", CURRENT_TIMESTAMP, '" & strUsuario & "','" & .TextMatrix(i, 12) & "') "
                            strSql = " INSERT INTO det_pedido(emp_codigo,dep_codigo, det_ped_cant_confirmada, " & _
                                 " det_ped_descripcion,det_ped_fechamod, det_ped_usumod," & _
                                 " ped_codigo,prd_codigo, det_ped_cant_pedida," & _
                                 " det_ped_cant_entregada, det_ped_precio,det_ped_dcto," & _
                                 " det_ped_incentivo) " & _
                                 " VALUES ('" & strEmpresa & "','PRI','0'," & _
                                 " '',CURRENT_TIMESTAMP,'" & strUsuario & "'," & _
                                 " '" & Pedido & "','" & clsConsulta.adorec_Def("prd_codigo_entregar") & "','" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") & "'," & _
                                 " '" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") & "','" & clsConsulta.adorec_Def("pro_pre_pre_prd_precio") & "','" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_dcto") & "', " & _
                                 " '0') "
                        Else
                            strSql = " UPDATE det_pedido " & _
                                     " SET det_ped_cant_pedida=det_ped_cant_pedida+" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") & "," & _
                                     " det_ped_cant_entregada=det_ped_cant_entregada+" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") & "," & _
                                     " det_ped_precio=" & (clsEjecuta.adorec_Def("det_ped_precio") * clsEjecuta.adorec_Def("det_ped_cant_pedida") + NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") * clsConsulta.adorec_Def("pro_pre_pre_prd_precio")) / (clsEjecuta.adorec_Def("det_ped_cant_pedida") + NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad")) & "," & _
                                     " det_ped_dcto=det_ped_dcto+" & NumeroDePromo * clsConsulta.adorec_Def("pro_pre_pre_prd_cantidad") * clsConsulta.adorec_Def("pro_pre_pre_prd_dcto") & " " & _
                                     " WHERE emp_codigo='" & strEmpresa & "' " & _
                                     " AND ped_codigo=" & Pedido & " " & _
                                     " AND prd_codigo='" & clsConsulta.adorec_Def("prd_codigo_entregar") & "' " & _
                                     " AND dep_codigo='PRI' "
                        End If
                        clsEjecuta.Ejecutar strSql, "M"
                        booSeModifico = True
                        NumeroDePromo = 0
                    End If
                    PromoPrendaPrecio = PromoPrendaPrecio + FormatoD4(Dcto)
                    clsConsulta.adorec_Def.MoveNext
                Wend
            End If
        End If
        clsPromo.adorec_Def.MoveNext
    Wend
    If booSeModifico = True Then
        clsPedido.RecalculoTotal (Pedido)
    End If
End Function


Public Sub Conectar()
    Set AdoConn = New ADODB.Connection
    Set AdoConnMaster = New ADODB.Connection
    strUsuario = UCase(strUsuario)
    Dim aux As Long
    Dim aux1 As Double
    ' Cadena de conexión a la base de datos, esta esta para el uso de MyODBC
    AdoConn.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDLocal
    If strPuertoLocal <> "" Then
        AdoConn.ConnectionString = AdoConn.ConnectionString & ", " & strPuertoLocal
    End If
    AdoConnMaster.ConnectionString = "Provider=SQLOLEDB.1;" _
                           & "Persist Security Info=False;" _
                           & "User ID=" & strUsuario & ";" _
                           & "Password=" & strClave & ";" _
                           & "Initial Catalog=" & strBDD & ";" _
                           & "Data Source=" & strServidorBDDMaster
    If strPuertoMaster <> "" Then
        AdoConnMaster.ConnectionString = AdoConnMaster.ConnectionString & ", " & strPuertoMaster
    End If
    AdoConn.ConnectionTimeout = 120
    AdoConn.CommandTimeout = 120
    AdoConn.CursorLocation = adUseClient
    AdoConn.Open
    
    AdoConnMaster.ConnectionTimeout = 120
    AdoConnMaster.CommandTimeout = 120
    AdoConnMaster.CursorLocation = adUseClient
    AdoConnMaster.Open
End Sub

Public Sub DesConectar()
    AdoConn.Close
    AdoConnMaster.Close
    Set AdoConn = Nothing
    Set AdoConnMaster = Nothing
End Sub


Public Sub LeerImpresoras()
    Set reg = CreateObject("WScript.Shell")
    'Dim reg As New WshShell
    On Error GoTo ErrorNombreImpresoraTicket
    ImpresoraTicket = reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraTicket")
    GoTo NombreImpresoraEtiqueta
ErrorNombreImpresoraTicket:
    On Error GoTo -1
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraTicket", ""
NombreImpresoraEtiqueta:
    On Error GoTo ErrorNombreImpresoraEtiqueta
    ImpresoraEtiqueta = reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraEtiqueta")
    GoTo NombreImpresoraPorDefecto
ErrorNombreImpresoraEtiqueta:
    On Error GoTo -1
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraEtiqueta", ""
NombreImpresoraPorDefecto:
    On Error GoTo ErrorNombreImpresoraPorDefecto
    ImpresoraPorDefecto = reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraPorDefecto")
    Exit Sub
ErrorNombreImpresoraPorDefecto:
    On Error GoTo -1
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraPorDefecto", Printer.DeviceName

    If ImpresoraPorDefecto <> Printer.DeviceName Then
        If MsgBox("La impresora por defecto ha cambiado a " & Printer.DeviceName & vbNewLine & _
                  "Desea cambiar a la anterior(" & ImpresoraPorDefecto & ")?", vbYesNo, "Impresora por Defecto") = vbYes Then
            DefinirImpresoraPorDefecto ImpresoraPorDefecto
        Else
            ImpresoraPorDefecto = Printer.DeviceName
            GuardarImpresoras
        End If
    End If
End Sub

Public Sub GuardarImpresoras()
    Set reg = CreateObject("WScript.Shell")
    'Dim reg As New WshShell
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraTicket", ImpresoraTicket
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraEtiqueta", ImpresoraEtiqueta
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Impresoras\NombreImpresoraPorDefecto", ImpresoraPorDefecto

End Sub


Public Sub LeerBalanza()
    Set reg = CreateObject("WScript.Shell")
    'Dim reg As New WshShell
    On Error GoTo ErrorPuerto
    PuertoBalanza = reg.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\\Balanza\Puerto")
    Exit Sub
ErrorPuerto:
    PuertoBalanza = 1
    GuardarPuertoBalanza
End Sub

Public Sub GuardarPuertoBalanza()
    Set reg = CreateObject("WScript.Shell")
    'Dim reg As New WshShell
    reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\NEEDNEW\Balanza\Puerto", PuertoBalanza
End Sub

Public Sub DefinirImpresoraPorDefecto(strImpresor As String)
    Dim p As VB.Printer
    For Each p In VB.Printers
       If p.DeviceName = strImpresor Then
          Set VB.Printer = p
          Exit For
       End If
    Next
End Sub



Public Function revisarEmail(strEMail As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim esMail As New vbSendMail.clsSendMail
    Dim TestArray() As String
    revisarEmail = True
        TestArray = Split(strEMail, ";")
        For j = 0 To UBound(TestArray)
        If esMail.IsValidEmailAddress(TestArray(j)) = False Then
            revisarEmail = False
        End If
        Next j
End Function

Public Function ComboNegocioDataSource() As ADODB.Recordset
    Dim clsCom As New clsConsulta
    clsCom.Inicializar AdoConn, AdoConnMaster
    strSql = " SELECT tip_ped_codigo " & _
             " FROM usuario_negocio " & _
             " WHERE usu_codigo='" & strUsuario & "'"
    
    clsCom.Ejecutar (strSql)
    
    strSql = " SELECT tipo_pedido.tip_ped_codigo,tip_ped_nombre " & _
             " FROM tipo_pedido "
    If clsCom.adorec_Def.RecordCount > 0 Then
        strSql = strSql & " WHERE tipo_pedido.tip_ped_codigo IN (" & _
                 " SELECT tip_ped_codigo " & _
                 " FROM usuario_negocio " & _
                 " WHERE usu_codigo='" & strUsuario & "') "
    End If
    strSql = strSql & " ORDER BY tip_ped_nombre"
    
    clsCom.Ejecutar strSql
    Set ComboNegocioDataSource = clsCom.adorec_Def
End Function

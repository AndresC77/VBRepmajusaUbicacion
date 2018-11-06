VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public clsCon_Def As clsConsulta
Public strDoc As String
Public strFecha As String
Public strPersona As String

Public Sub Inicializar(ByVal adocon_ParL As ADODB.Connection, ByVal adocon_ParM As ADODB.Connection)
    Set clsCon_Def = New clsConsulta
    clsCon_Def.Inicializar adocon_ParL, adocon_ParM
End Sub

Private Sub Class_Terminate()
    Set clsCon_Def = Nothing
End Sub


Public Sub NuevoDet(Mot As String, prdIng As String, CantIng As Double, prdPed As String)
    Dim strSQL As String
    If prdIng <> "" And CantIng <> 0 Then
        strSQL = " SELECT count(*) as n " & _
                 " FROM det_cambio " & _
                 " WHERE emp_codigo='" & strEmpresa & "'" & _
                 " AND cam_codigo='" & strDoc & "'" & _
                 " AND mot_aju_codigo='" & Mot & "'" & _
                 " AND prd_codigo_ing='" & prdIng & "'" & _
                 " AND prd_codigo_ped='" & prdPed & "'"
        clsCon_Def.Ejecutar strSQL, "M"
        If clsCon_Def.adorec_Def("n") > 0 Then
            strSQL = " UPDATE det_cambio " & _
                     " SET det_cam_cantidad=det_cam_cantidad+'" & CantIng & "'" & _
                     " WHERE emp_codigo='" & strEmpresa & "'" & _
                     " AND cam_codigo='" & strDoc & "'" & _
                     " AND mot_aju_codigo='" & Mot & "'" & _
                     " AND prd_codigo_ing='" & prdIng & "'" & _
                     " AND prd_codigo_ped='" & prdPed & "'"
        Else
            strSQL = " INSERT INTO det_cambio(emp_codigo,cam_codigo," & _
                     " mot_aju_codigo,prd_codigo_ing,det_cam_cantidad," & _
                     " prd_codigo_ped,tip_ing_codigo,ing_codigo,det_cam_fechamod,det_cam_usumod) " & _
                     " VALUES('" & strEmpresa & "','" & strDoc & "'," & _
                     " '" & Mot & "','" & prdIng & "','" & CantIng & "'," & _
                     " '" & prdPed & "','','0',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        End If
        clsCon_Def.Ejecutar strSQL, "M"
    End If
End Sub

Public Function NuevoCambio(ConfirmaNumero As Boolean, Optional strSuc As String = "", Optional strPto As String = "", Optional Persona As String = "", Optional Fecha As String = "", Optional Factura As String = "", Optional Observ As String = "") As Boolean
    Dim strSQL As String
    Dim Doc As String
    Observ = UCase(Observ)
    If Fecha = "" Then
        strFecha = Format(HoyDia, "yyyy-mm-dd")
    Else
        strFecha = Fecha
    End If
    NuevoCambio = False
    If strSuc = "" And strPto = "" Then
        strSuc = strSucursal
        strPto = strPtoFactura
    End If
    
    If NumeroDeNuevoCambio(ConfirmaNumero, strSuc, strPto, Doc) = True Then
        strDoc = strSuc & strPto & Format(Doc, "0000000")
        strSQL = " INSERT INTO cambio (cam_codigo,emp_codigo," & _
                 " per_codigo,cam_fecha," & _
                 " cam_observacion,cam_factura,cam_fechamod,cam_usumod) VALUES(" & _
                 " '" & strDoc & "','" & strEmpresa & "'," & _
                 " '" & Persona & "','" & strFecha & "'," & _
                 " '" & Observ & "','" & Factura & "',CURRENT_TIMESTAMP,'" & strUsuario & "')"
        clsCon_Def.Ejecutar strSQL, "M"
        strPersona = Persona
        NuevoCambio = True
    End If
End Function

Private Function NumeroDeNuevoCambio(ConfirmaNumero As Boolean, strSuc As String, strPto As String, ByRef Doc As String) As Boolean
    Dim strSQL As String
    Dim NumCambio As Double
    Dim booPasar As Boolean
    Dim strDocume As String
    strSQL = " SELECT COALESCE(RIGHT(MAX(cam_codigo+1),7),1) AS num " & _
             " FROM cambio " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cam_codigo like '" & FormatoD0(strSuc & strPto) & "%' " & _
             " GROUP BY emp_codigo"
    clsCon_Def.Ejecutar strSQL
    If clsCon_Def.adorec_Def.RecordCount > 0 Then
        Doc = FormatoD0(clsCon_Def.adorec_Def("num"))
    Else
        Doc = FormatoD0(Format(1, "0000000"))
    End If
    If ConfirmaNumero = True Then
        booPasar = False
        NumeroDeNuevoCambio = True
        While booPasar = False
            Doc = Val(InputBox("CAMBIO A GENERAR", "Numeración de Documento", Doc))
            
            strDoc = strSuc & strPto & Format(Doc, "0000000")
            If strDoc = strSuc & strPto & Format(0, "0000000") Then
                booPasar = True
                NumeroDeNuevoCambio = False
            Else
                strSQL = " SELECT IIF(count(*) is null or count(*)=0,'N','S') as sn " & _
                         " FROM cambio " & _
                         " WHERE emp_codigo='" & strEmpresa & "' " & _
                         " AND cam_codigo='" & strDoc & "'"
                clsCon_Def.Ejecutar (strSQL)
                If clsCon_Def.adorec_Def("sn") = "S" Then
                    If MsgBox("CAMBIO YA INGRESADO" & vbNewLine & "Compruebe la numeración." & _
                    vbNewLine & "SI, INGRESAR NUEVO NUMERO" & vbNewLine & _
                    "NO, PASAR Y NO INGRESAR", vbCritical + vbYesNo, "Numeración de Documento") = vbNo Then
                        booPasar = True
                        NumeroDeNuevoCambio = False
                    End If
                Else
                    booPasar = True
                End If
            End If
        Wend
    Else
        NumeroDeNuevoCambio = True
    End If
End Function

Public Sub AsignarIngreso(strCambio As String, Mot As String, prdIng As String, prdPed As String, TipoIng As String, NumeroIng As String)
    Dim strSQL As String
    strSQL = " UPDATE det_cambio " & _
             " SET tip_ing_codigo='" & TipoIng & "', " & _
             " ing_codigo='" & NumeroIng & "', " & _
             " det_cam_usumod='" & strUsuario & "', " & _
             " det_cam_fechamod=CURRENT_TIMESTAMP " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND cam_codigo='" & strCambio & "' " & _
             " AND mot_aju_codigo='" & Mot & "' " & _
             " AND prd_codigo_ing='" & prdIng & "' " & _
             " AND prd_codigo_ped='" & prdPed & "'"
    clsCon_Def.Ejecutar strSQL, "M"
End Sub

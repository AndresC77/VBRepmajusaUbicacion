VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'################################################################################'
'#  clase clsConsulta v.01                                                      #
'#  Copyright (C) 2002                                                          #
'#                                                                              #
'#  Clase para ejecutar cualquier SQL contra la base de datos y obtener un      #
'#  ADODB.Recordset con los resultados esto sin necesidad de abrir y            #
'#  cerrar el Recordset varias veces.                                           #
'#  Atributos:                                                                  #
'#  adocomm_Def ADODB.Command Objeto PRIVADO que ejecuta las sentencias SQL     #
'#  adorec_Def ADODB.Recordset Objeto PUBLICO que guarda los resultados de      #
'#                             la consulta.                                     #
'#  adocon_Def AS ADODB.Connection Objeto PRIVADO que mantine la conexion       #
'#                                 al motor de base de datos.                   #
'#  Procedimientos:                                                             #
'#  Inicializar(ADODB.Connection, strSQL) Carga la connexión y el SQL para la   #
'#                                        consulta.                             #
'#  Ejecutar(strSQL) Ejecuta la consulta y almacena los resultados en el        #
'#                   Recordset adorec_Def                                       #
'#  Actualizar() Actualiza el Recordset volviendo a consultar a la base         #
'#               de datos                                                       #
'#                                                                              #
'#  Filtrar(strFiltro) Filtra al recordset (adorec_Def). en strFiltro debe      #
'#                     llegar un WHERE de SQL                                   #
'#                                                                              #
'#  QuitarFiltro() Quita el filtro al recordset (adorec_Def)                    #
'#                                                                              #
'################################################################################
'/****************************************************************************/'

Private adocomm_DefM As ADODB.Command
Private adocomm_DefL As ADODB.Command
Public adorec_Def As ADODB.Recordset
Private adocon_DefM As ADODB.Connection
Private adocon_DefL As ADODB.Connection
Private Don As String
Public Ejecuto As Boolean

Public Sub Inicializar(adocon_ParL As ADODB.Connection, adocon_ParM As ADODB.Connection)
    Set adocomm_DefM = New ADODB.Command
    Set adocomm_DefL = New ADODB.Command
    Set adocon_DefM = adocon_ParM
    Set adocon_DefL = adocon_ParL
    adocomm_DefM.CommandText = strSQ
    adocomm_DefL.CommandText = strSQ
    adocomm_DefM.ActiveConnection = adocon_DefM
    adocomm_DefL.ActiveConnection = adocon_DefL
End Sub
    
Public Sub Ejecutar(Optional strSQ As Variant = "", Optional donde As String = "L", Optional ConError As Boolean = True)
    Don = donde
    Ejecuto = True
    Dim i As Integer
    Dim j As Integer
    j = 0
    i = 0
Ejecutar:
    On Error GoTo -1
    If donde = "L" Then
        If strSQ <> "" Then
            adocomm_DefL.CommandText = strSQ
        End If
        On Error GoTo errhandler
            Set adorec_Def = adocomm_DefL.Execute
    Else
        If strSQ <> "" Then
            adocomm_DefM.CommandText = strSQ
        End If
        On Error GoTo errhandler
            Set adorec_Def = adocomm_DefM.Execute
            'ChequeoSincronizacion
    End If
    Exit Sub
errhandler:
    Ejecuto = False
    If ConError = True Then
        Select Case Err.Number
            Case 1046
                MsgBox " When you perform a normal mysql_connect and " & vbCrLf & _
                       " not a mysql_real_connect you have to choose a " & vbCrLf & _
                       " database, so Please Choose a database."
            Case 1205
                GoTo Ejecutar
            Case -2147217871
                If j < 3 Then
                    j = j + 1
                    GoTo Ejecutar
                Else
                    MsgBox "[" & Err.Number & "] " & Err.Description & vbNewLine & vbNewLine & strSQ
                    If MsgBox("Comuniquese con el administrador del NEED" & vbNewLine & "Desea volver a ejecutar?", vbYesNo + vbQuestion + vbDefaultButton1, "Ejecucion BDD") = vbYes Then
                        GoTo Ejecutar
                    End If
                End If
            Case -2147467259
                If i < 20 Then
                    
'                    DesConectar
'                    Conectar
'                    Inicializar AdoConn, AdoConnMaster
                    i = i + 1
                    GoTo Ejecutar
                Else
                    MsgBox "[" & Err.Number & "] " & Err.Description & vbNewLine & vbNewLine & strSQ
                End If
            Case Else
                If Right(Err.Description, 50) = "' was locked with a READ lock and can't be updated" _
                    Or Err.Number = -2147217871 _
                    Or InStr(1, UCase(Err.Description), "RERUN") > 0 Then
                    MsgBox " Se va a volver a intentar ejecutar la consulta:" & vbNewLine & strSQ
                    GoTo Ejecutar
                Else
                    MsgBox "[" & Err.Number & "] " & Err.Description & vbNewLine & vbNewLine & strSQ
                End If
        
        End Select
        'MsgBox strSQ
    End If
End Sub

Private Sub ChequeoSincronizacion()
    Dim strMaster As String
    Dim strSlave As String
    Dim strMasterOld As String
    Dim strSlaveOld As String
    Dim Sale As Boolean
    Dim i As Long
    Dim adorec_DefM As New ADODB.Recordset
    Dim adorec_DefL As New ADODB.Recordset
    
    strMaster = "SHOW MASTER STATUS"
    strSlave = "SHOW SLAVE STATUS"
    strMasterOld = adocomm_DefM.CommandText
    strSlaveOld = adocomm_DefL.CommandText
    adocomm_DefM.CommandText = strMaster
    adocomm_DefL.CommandText = strSlave
    Sale = False
    i = 0
    While Sale = False
        Set adorec_DefM = adocomm_DefM.Execute
        Set adorec_DefL = adocomm_DefL.Execute
        If adorec_DefM(0) = adorec_DefL(4) And adorec_DefM(1) = adorec_DefL(5) And adorec_DefM(1) = adorec_DefL(16) And adorec_DefL(9) = "Yes" And adorec_DefL(10) = "Yes" Then
            Sale = True
        End If
        i = i + 1
        If i = 1000 Then
            MsgBox "Debe segir esperando. Comunique al administrador. No cierre el sistema" & vbNewLine & adorec_DefM(0) & " = " & adorec_DefL(4) & " And " & adorec_DefM(1) & " = " & adorec_DefL(5) & " And " & adorec_DefM(1) & " = " & adorec_DefL(16) & " And " & adorec_DefL(9) & " = Yes And " & adorec_DefL(10) & " = Yes"
            i = 0
        End If
    Wend
    adocomm_DefM.CommandText = strMasterOld
    adocomm_DefL.CommandText = strSlaveOld
End Sub

Public Sub Actualizar()
    If Don = "L" Then
        Set adorec_Def = adocomm_DefL.Execute
    Else
        Set adorec_Def = adocomm_DefM.Execute
    End If
End Sub

Public Sub Filtrar(strFiltro As String)
   ' Establece un filtro sobre el objeto Recordset especificado
   QuitarFiltro
   adorec_Def.Filter = strFiltro
End Sub

Public Sub QuitarFiltro()
   ' Quita filtro sobre el objeto Recordset especificado
   adorec_Def.Filter = adFilterNone
End Sub
Private Sub Class_Terminate()
    Set adorec_Def = Nothing
    Set adocomm_DefM = Nothing
    Set adocon_DefM = Nothing
    Set adocomm_DefL = Nothing
    Set adocon_DefL = Nothing
End Sub


Public Function EjecutarSeguro(Optional strSQ As String = "") As Boolean
    If strSQ <> "" Then
        adocomm_DefM.CommandText = strSQ
    End If
    On Error GoTo errhandler
    Set adorec_Def = adocomm_Def.Execute
    ErroresConsulta = 0
    EjecutarSeguro = True
    Exit Function
    
errhandler:
    ErroresConsulta = ErroresConsulta + 1
    '100 intentos de hacer la operación, han de ser suficientes
    If ErroresConsulta = 100 Then
        MsgBox "Ha habido un error en una operación de la base de datos. Se intentó 100 veces realizar la operación.", vbCritical, "Error"
        EjecutarSeguro = True
    Else
        EjecutarSeguro = False
    End If
End Function

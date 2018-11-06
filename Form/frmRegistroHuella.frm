VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRegistroHuella 
   BackColor       =   &H00DDDDDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Huella"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmRegistroHuella.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6990
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDDD&
      Caption         =   "Empleados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin MSDataListLib.DataCombo cmbEmpleado 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   375
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione un Empleado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   270
         TabIndex        =   10
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4312
      TabIndex        =   7
      Top             =   4320
      Width           =   1500
   End
   Begin VB.CommandButton btnLimpiar 
      Caption         =   "&Limpiar"
      Height          =   360
      Left            =   2632
      TabIndex        =   6
      Top             =   4320
      Width           =   1500
   End
   Begin VB.CommandButton btnAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   952
      TabIndex        =   5
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Frame fraContenedor 
      BackColor       =   &H00DDDDDD&
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6735
      Begin VB.Frame fraHuella 
         BackColor       =   &H00DDDDDD&
         Caption         =   "Huella"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   6495
         Begin VB.CommandButton btnAgregarHuella 
            Height          =   975
            Left            =   120
            Picture         =   "frmRegistroHuella.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Agregar"
            Top             =   360
            Width           =   975
         End
         Begin VB.Frame fraExtra 
            BackColor       =   &H00DDDDDD&
            Height          =   1215
            Left            =   4800
            TabIndex        =   12
            Top             =   200
            Width           =   1455
            Begin VB.CommandButton btnVerificarHuella 
               Caption         =   "&Verificar"
               Height          =   345
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton btnBorrarHuella 
               Caption         =   "&Borrar"
               Height          =   345
               Left            =   240
               TabIndex        =   13
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Label lblInfoHuella 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            Height          =   915
            Left            =   1200
            TabIndex        =   16
            Top             =   360
            Width           =   3405
         End
      End
      Begin VB.TextBox txtConfirmaClave 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2640
         MaxLength       =   3
         PasswordChar    =   "l"
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2640
         MaxLength       =   3
         PasswordChar    =   "l"
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1545
         TabIndex        =   9
         Top             =   465
         Width           =   855
      End
      Begin VB.Label lblConfirmaClave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirma Contraseña:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   885
         TabIndex        =   8
         Top             =   825
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmRegistroHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsSql As New clsConsulta
Private clsAux As New clsConsulta
Private strSql As String
Private Registrado As Boolean
Private Mensaje As String
Private codigo As String
Private Verificar As String
Private foto As String
Private Actualizar As Boolean
Private Interno As Boolean
Private conSlave As ADODB.Connection
Private conMaster As ADODB.Connection
Private BaseDatos As String

Private Sub CargaEmpleados()
    strSql = " SELECT CONCAT(epl_apellidos,' ',epl_nombres) as nombre,epl_codigo as codigo " & _
             " FROM empleado " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " ORDER BY 1 "
    clsSql.Ejecutar strSql
    Set cmbEmpleado.RowSource = clsSql.adorec_Def.DataSource
    cmbEmpleado.ListField = "nombre"
    cmbEmpleado.BoundColumn = "codigo"
End Sub

Private Sub btnAgregar_Click()
    On Error Resume Next
    If btnAgregar.Caption = "&Aceptar" Then
        If ComprobarDatos(0) Then
            If Registrado = True Then
                Insertar
                MsgBox "Se ha registrado el Empleado " & cmbEmpleado.Text & " exitosamente!", vbInformation, "Registro de Huella"
                Limpiar
                cmbEmpleado.SetFocus
            Else
                If ComprobarClave Then
                    If MsgBox("Desea registrar el Empleado sin huella", vbQuestion + vbYesNo, "Registro de Huella") = vbYes Then
                        If GenerarClave Then
                            Insertar
                            MsgBox "Se ha registrado el Empleado " & cmbEmpleado.Text & " exitosamente!", vbInformation, "Registro de Huella"
                            Limpiar
                            cmbEmpleado.SetFocus
                        Else
                            MsgBox "Ingrese otra contraseña porque ya existe una igual", vbCritical, "Contraseña"
                        End If
                    End If
                End If
            End If
        End If
    Else
        'Se actualizara el profesor
        If ComprobarDatos(0) Then
            If Registrado = True Then
                Modificar
                Mensaje = "Se ha modificado el Empleado " & cmbEmpleado.Text & " exitosamente!"
                MsgBox Mensaje, vbInformation, "Registro de Huella"
                Limpiar
                cmbEmpleado.SetFocus
            Else
                If ComprobarClave Then
                    If MsgBox("Desea registrar el Empleado sin huella digital?", vbQuestion + vbYesNo, "Registro de Huella") = vbYes Then
                        Modificar
                        Mensaje = "Se ha registrado el Empleado " & cmbEmpleado.Text & " exitosamente!"
                        MsgBox Mensaje, vbInformation, "Registro de Huella"
                        Limpiar
                        cmbEmpleado.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub btnAgregarHuella_Click()
    Dim verifico As Boolean
    verifico = False
    
    If ComprobarDatos(1) Then
        If ComprobarClave Then
            If GenerarClave Then
            
                CargarHuella
                ATok = 0
                'If ATok = ATInit() Then
                 '   If ATok = ATOpenSensorW(0, 1) Then
                        num = FreeFile
                        archivoaux = Path & "\NuevoDocente.dat"
                        Open archivoaux For Output As #num
                            Print #num, codigo;
                        Close #num
                        '0 Enroll  1 Verify  Codigo delete
                        num = FreeFile
                        Archivo = Path & "\Proceso.dat"
                        Open Archivo For Output As #num
                            Print #num, "0";
                        Close #num
                        
                        'enroll = ShellExecute(0, "open", "C:\TESIS DOCS\AuthenTec\ATSC\ATSource\ATWinLIB\ATDemoCode\ATLowLevelApiDemoEnroll\Debug\ATLowLevelApiDemo.exe", "", "", 3)
                        enroll = Shell(PathCpp, vbNormalFocus)
                        If enroll = 0 Then
                            MsgBox "Ocurrió un error inesperado, intente más tarde", vbCritical, "Huella digital"
                        Else
                             hProcess = OpenProcess(&H100000, True, enroll)
                             WaitForSingleObject hProcess, -1
                             CloseHandle hProcess
                        End If
                        If Dir(Archivo) <> "" Then Kill (Archivo) 'elimino proceso.txt
                        If Dir(archivoaux) <> "" Then Kill (archivoaux)  'elimino nuevodocente.txt
                        num = FreeFile
                        Archivo = Path & "\Comprobar.dat"
                        
                        Open Archivo For Binary As #num
                            nLen = LOF(num)
                            texto = Space(nLen)
                        Get #num, , texto
                        Close #num
                        If Dir(Archivo) <> "" Then Kill (Archivo)
                        
                        If FormatoD0(texto) > -1 Then
                            Registrado = Bool(texto)
                            
                            If Registrado Then
                                MsgBox "Se ha registrado la huella del Empleado " & cmbEmpleado.Text & " exitosamente", vbInformation, "Huella Digital"
                                btnAgregarHuella.Enabled = False
                                'btnBorrarHuella.Enabled = True
                                'btnVerificarHuella.Enabled = True
                                   
                                   
                                GuardaHuella
                                   
                                If MsgBox("Desea verificar la huella registrada?", vbQuestion + vbYesNo, "Verificar Huella Digital") = vbYes Then
                                    While verifico = False
                                        
                                        'ATok = ATCloseSensor()
                                        'ATok = ATClose()
                                        'ATok = 0
                                        'If ATok = ATInit() Then
                                         '   If ATok = ATOpenSensorW(0, 1) Then
                                                num = FreeFile
                                                Archivo = Path & "\Proceso.dat"
                                                Open Archivo For Output As #num
                                                    Print #num, "2";
                                                Close #num
                                                enroll = Shell(PathCpp, vbNormalFocus)
                                                If enroll = 0 Then
                                                 '
                                                 'Handle Error, Shell Didn't Work
                                                 '
                                                Else
                                                     hProcess = OpenProcess(&H100000, True, enroll)
                                                     WaitForSingleObject hProcess, -1
                                                     CloseHandle hProcess
                                                     'Me.Show
                                                End If
                                                If Dir(Archivo) <> "" Then Kill (Archivo)
                                                
                                                num = FreeFile
                                                Archivo = Path & "\Verificar.dat"
                                                
                                                Open Archivo For Binary As #num
                                                    nLen = LOF(num)
                                                    texto = Space(nLen)
                                                Get #num, , texto
                                                Close #num
                                                Verificar = CStr(texto)
                                                If Dir(Archivo) <> "" Then Kill (Archivo)
                                                
                                                num = FreeFile
                                                Archivo = Path & "\Comprobar.dat"
                                                
                                                Open Archivo For Binary As #num
                                                    nLen = LOF(num)
                                                    texto = Space(nLen)
                                                Get #num, , texto
                                                Close #num
                                                
                                                If Dir(Archivo) <> "" Then Kill (Archivo)
                                                If FormatoD0(texto) > -1 Then
                                                    If codigo <> Verificar Then
                                                       MsgBox "La huella no es reconocida", vbInformation, "Verificar Huella Digital"
                                                       If MsgBox("Desea reintentar la verificación de huella?", vbQuestion + vbYesNo, "Verificar Huella Digital") = vbNo Then verifico = True
                                                    Else
                                                       MsgBox "Se verificó la huella. Es " & cmbEmpleado.Text, vbInformation, "Verificar Huella Digital"
                                                       verifico = True
                                                    End If
                                                Else
                                                    verifico = True
                                                End If
                                          '  Else
                                          '      MsgBox "Error: El dispositivo no esta conectado", vbInformation, "APC Biopod"
                                          '  End If
                                        'Else
                                        '    MsgBox "Error: El controlador del dispositivo no esta correctamente instalado", vbInformation, "APC Biopod"
                                        'End If
                                    Wend
                                End If
                                btnBorrarHuella.Enabled = True
                                btnVerificarHuella.Enabled = True
                            Else
                                MsgBox "No se a registrado la Huella Digital", vbInformation, "Huella Digital"
                                btnAgregarHuella.Enabled = True
                                btnBorrarHuella.Enabled = False
                                btnVerificarHuella.Enabled = False
                            End If
                        Else
                            Registrado = False
                        End If
                    'Else
                    '   MsgBox "Error: El dispositivo no esta conectado", vbInformation, "APC Biopod"
                   ' End If
                'Else
                 '   MsgBox "Error: El controlador del dispositivo no esta correctamente instalado", vbInformation, "APC Biopod"
                'End If
                'ATok = ATCloseSensor()
                'ATok = ATClose()
            Else
                MsgBox "Ingrese otra contraseña porque ya existe una igual", vbCritical, "Contraseña"
            End If
        End If
    End If
    
End Sub

Private Sub btnBorrarHuella_Click()
    If ComprobarClave Then
        If ComprobarDatos(1) Then
            If GenerarClave Then
                If MsgBox("Desea eliminar las huellas digitales registradas de " & cmbEmpleado.Text & "?" & vbCrLf & "Si acepta, tendrá que ingresar nuevamente la huella del empleado", vbQuestion + vbYesNo, "Eliminar Huella Digital") = vbYes Then
                    ATok = 0
                    CargarHuella
                    'If ATok = ATInit() Then
                     '   If ATok = ATOpenSensorW(0, 1) Then
                            num = FreeFile
                            Archivo = Path & "\Proceso.dat"
                            Open Archivo For Output As #num
                                Print #num, codigo;
                            Close #num
                            enroll = Shell(PathCpp, vbNormalFocus)
                            If enroll = 0 Then
                             '
                             'Handle Error, Shell Didn't Work
                             '
                            Else
                                 hProcess = OpenProcess(&H100000, True, enroll)
                                 WaitForSingleObject hProcess, -1
                                 CloseHandle hProcess
                                 'Me.Show
                            End If
                            If Dir(Archivo) <> "" Then Kill (Archivo)
                            BorrarHuella cmbEmpleado.BoundText
                            MsgBox "Se eliminó las huellas exitosamente!", vbInformation, "Eliminar Huella Digital"
                            btnBorrarHuella.Enabled = False
                            btnVerificarHuella.Enabled = False
                            btnAgregarHuella.Enabled = True
                           GuardaHuella
                        'Else
                        '    MsgBox "Error: El dispositivo no esta conectado", vbInformation, "APC Biopod"
                        'End If
                    'Else
                    '    MsgBox "Error: El controlador del dispositivo no esta correctamente instalado", vbInformation, "APC Biopod"
                    'End If
                    'ATok = ATCloseSensor()
                    'ATok = ATClose()
                End If
            Else
                MsgBox "Ingrese otra contraseña porque ya existe una igual", vbCritical, "Contraseña"
            End If
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub


Private Sub btnLimpiar_Click()
    If btnAgregar.Caption = "&Modificar" Then
        If Registrado = True Then
            If Interno = True Then
                strSql = " SELECT COALESCE(epl_huella_interno,0) as huella " & _
                         " FROM empleado_huella " & _
                         " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                         " AND emp_codigo='" & strEmpresa & "' "
                clsSql.Ejecutar strSql
            Else
                strSql = " SELECT COALESCE(epl_huella_base,0) as huella " & _
                         " FROM empleado_huella " & _
                         " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                         " AND emp_codigo='" & strEmpresa & "' "
                clsSql.Ejecutar strSql
            End If
            If clsSql.adorec_Def.RecordCount > 0 Then
                If Abs(FormatoD0(clsSql.adorec_Def("huella"))) = 0 Then
                    ATok = 0
                    If MsgBox("Desea cancelar todos los cambios realizados?", vbQuestion + vbYesNo, "Registro Docente") = vbYes Then
                    
                    'If ATok = ATInit() Then
                     '   If ATok = ATOpenSensorW(0, 1) Then
                            num = FreeFile
                            Archivo = Path & "\Proceso.dat"
                            Open Archivo For Output As #num
                                Print #num, codigo;
                            Close #num
                            enroll = Shell(PathCpp, vbNormalFocus)
                            If enroll = 0 Then
                             '
                             'Handle Error, Shell Didn't Work
                             '
                            Else
                                 hProcess = OpenProcess(&H100000, True, enroll)
                                 WaitForSingleObject hProcess, -1
                                 CloseHandle hProcess
                                 'Me.Show
                            End If
                            If Dir(Archivo) <> "" Then Kill (Archivo)
                            GuardaHuella
                    Else
                        Exit Sub
                    End If
                            'BorrarHuella codigo
                            'MsgBox "Se eliminó las huellas exitosamente", vbInformation, "Huella Digital"
                            'btnBorrarHuella.Enabled = False
                            'btnVerificarHuella.Enabled = False
                            'btnAgregarHuella.Enabled = True
                      '  Else
                       '     MsgBox "Error: El dispositivo no esta conectado", vbInformation, "APC Biopod"
                       ' End If
                    'Else
                    '    MsgBox "Error: El controlador del dispositivo no esta correctamente instalado", vbInformation, "APC Biopod"
                    'End If
                    'ATok = ATCloseSensor()
                    'ATok = ATClose()
                        
                        
                End If
            End If
        End If
    End If
    Limpiar
    cmbEmpleado.SetFocus
End Sub

Private Sub btnVerificarHuella_Click()
    If ComprobarDatos(1) Then
        If GenerarClave Then
            CargarHuella
            ATok = 0
            'If ATok = ATInit() Then
                'If ATok = ATOpenSensorW(0, 1) Then
                    num = FreeFile
                    Archivo = Path & "\Proceso.dat"
                    Open Archivo For Output As #num
                        Print #num, "1";
                    Close #num
                    enroll = Shell(PathCpp, vbNormalFocus)
                    If enroll = 0 Then
                     '
                     'Handle Error, Shell Didn't Work
                     '
                    Else
                         hProcess = OpenProcess(&H100000, True, enroll)
                         WaitForSingleObject hProcess, -1
                         CloseHandle hProcess
                         'Me.Show
                    End If
                    If Dir(Archivo) <> "" Then Kill (Archivo)
                    num = FreeFile
                    Archivo = Path & "\Verificar.dat"
                    
                    Open Archivo For Binary As #num
                        nLen = LOF(num)
                        texto = Space(nLen)
                    Get #num, , texto
                    Close #num
                    If Dir(Archivo) <> "" Then Kill (Archivo)
                    Verificar = CStr(texto)
                    
                    num = FreeFile
                    Archivo = Path & "\Comprobar.dat"
                    
                    Open Archivo For Binary As #num
                        nLen = LOF(num)
                        texto = Space(nLen)
                    Get #num, , texto
                    Close #num
                    
                    If Dir(Archivo) <> "" Then Kill (Archivo)
                    If FormatoD0(texto) > -1 Then
                        VerificarDocente Verificar
                    End If
                    
                                        
                'Else
                '    MsgBox "Error: El dispositivo no esta conectado", vbInformation, "APC Biopod"
                'End If
            'Else
            '    MsgBox "Error: El controlador del dispositivo no esta correctamente instalado", vbInformation, "APC Biopod"
            'End If
            'ATok = ATCloseSensor()
            'ATok = ATClose()
        End If
    End If
    'MsgBox "Se eliminó las huellas exitosamente", vbInformation, "Huella Digital"
End Sub

Private Sub Form_Load()
    Me.Left = (mdiPrincipal.Width - Me.Width) / 2
    Me.Top = 0
    Interno = True
    BaseDatos = ""
    num = FreeFile
    Archivo = Path & "\Path.dat"

    If Dir(Archivo, vbHidden) <> "" Then SetAttr Archivo, vbNormal: Kill (Archivo)
    Open Archivo For Output As #num
        Print #num, Path;
    Close #num

    num = FreeFile
    Archivo = Path & "\Huellas.dat"

    If Dir(Archivo, vbHidden) <> "" Then SetAttr Archivo, vbNormal: Kill (Archivo)
    Open Archivo For Output As #num
        Print #num, Path & "\";
    Close #num

    strPathHuella = Path & "\"
    SetAttr Archivo, vbHidden
    
    clsSql.Inicializar AdoConn, AdoConnMaster
    
    strSql = " SELECT par_numero,par_texto " & _
             " FROM parametro " & _
             " WHERE emp_codigo='" & strEmpresa & "' " & _
             " AND par_codigo='REG' "
    clsSql.Ejecutar strSql
    
    If clsSql.adorec_Def.RecordCount > 0 Then
        Interno = CBool(clsSql.adorec_Def(0))
        BaseDatos = clsSql.adorec_Def(1)
    Else
        MsgBox "Ingrese parámetros de REGISTRO HUELLA", vbInformation, "Parametro"
        Unload Me
    End If
    
    If Interno = False Then
        Set conSlave = New ADODB.Connection
        Set conMaster = New ADODB.Connection
        
        conSlave.ConnectionString = "driver={MySQL ODBC 3.51 Driver};" _
                                 & "server=" & strServidorBDDLocal & ";" _
                                 & "uid=" & strUsuario & ";" _
                                 & "pwd=" & strClave & ";" _
                                 & "database=" & BaseDatos
        If strPuertoLocal <> "" Then
            conSlave.ConnectionString = conSlave.ConnectionString & ";port=" & strPuertoLocal
        End If
        
        conMaster.ConnectionString = "driver={MySQL ODBC 3.51 Driver};" _
                                 & "server=" & strServidorBDDMaster & ";" _
                                 & "uid=" & strUsuario & ";" _
                                 & "pwd=" & strClave & ";" _
                                 & "database=" & BaseDatos
        If strPuertoLocal <> "" Then
            conMaster.ConnectionString = conMaster.ConnectionString & ";port=" & strPuertoLocal
        End If
        
        conSlave.ConnectionTimeout = 30
        conSlave.CursorLocation = adUseClient
        conSlave.Open
        
        conMaster.ConnectionTimeout = 30
        conMaster.CursorLocation = adUseClient
        conMaster.Open
        
        clsAux.Inicializar conSlave, conMaster
        
        
    
    End If
    
    
    lblInfoHuella.Caption = "Para registrar la huella digital es necesario " & vbCrLf & _
                            "concectar el dispositivo APC biopod. Cuando " & vbCrLf & _
                            "haya registrado correctamente la huella no " & vbCrLf & _
                            "podrá cambiar la clave del empleado."
    

    Registrado = False
    CargaEmpleados
    Limpiar
End Sub


Private Sub Limpiar()
    Deshabilitar False
    Actualizar = False
    
    codigo = ""
    btnAgregar.Caption = "&Aceptar"
    
    cmbEmpleado.BoundText = ""
    
    txtClave.Text = ""
    txtConfirmaClave.Text = ""
    
    btnAgregarHuella.Enabled = False
    btnVerificarHuella.Enabled = False
    btnBorrarHuella.Enabled = False
    
    Registrado = False
    
End Sub

Private Sub Insertar()
    Dim codigo As String, clave As String
    Dim i As Long
    codigo = cmbEmpleado.BoundText
    clave = txtClave.Text
    
    If Interno = True Then
        strSql = " INSERT INTO empleado_huella(emp_codigo,epl_codigo,epl_huella_clave,epl_huella_registrado," & _
                 " epl_huella_fechamod,epl_huella_usumod,epl_huella_interno) " & _
                 " VALUES('" & strEmpresa & "','" & cmbEmpleado.BoundText & "','" & clave & "','" & IIf(Registrado = True, 1, 0) & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "','" & IIf(Registrado = True, 1, 0) & "')"
        clsSql.Ejecutar strSql, "M"
    Else
        strSql = " INSERT INTO empleado_huella(emp_codigo,epl_codigo,epl_huella_clave,epl_huella_registrado," & _
                 " epl_huella_fechamod,epl_huella_usumod,epl_huella_base) " & _
                 " VALUES('" & strEmpresa & "','" & cmbEmpleado.BoundText & "','" & clave & "','" & IIf(Registrado = True, 1, 0) & "'," & _
                 " CURRENT_TIMESTAMP,'" & strUsuario & "','" & IIf(Registrado = True, 1, 0) & " ')"
        clsSql.Ejecutar strSql, "M"
    End If
End Sub

Private Sub VerificarDocente(id As String)
     strSql = " SELECT empleado_huella.epl_codigo,CONCAT(epl_apellidos,' ',epl_nombres) as nomb " & _
              " FROM empleado_huella " & _
              " INNER JOIN empleado " & _
              " ON empleado.emp_codigo=empleado_huella.emp_codigo " & _
              " AND empleado.epl_codigo=empleado_huella.epl_codigo " & _
              " WHERE empleado_huella.epl_huella_clave='" & id & "' " & _
              " AND empleado_huella.emp_codigo='" & strEmpresa & "' "
     clsSql.Ejecutar strSql
     
     If clsSql.adorec_Def.RecordCount = 0 Then
        MsgBox "La huella no es reconocida", vbInformation, "Verificar Huella"
     Else
        MsgBox "Es el empleado " & clsSql.adorec_Def("nomb"), vbInformation, "Verificar Huella"
     End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To Me.Controls.count - 1
        Set Me.Controls(i).DataSource = Nothing
    Next i
    On Error GoTo 0
    Set clsSql = Nothing
    Set clsAux = Nothing
    If Interno = False Then
        conSlave.Close
        conMaster.Close
    End If
    
    codigo = ""
    Registrado = False
    Actualizar = False
    
    Archivo = Path & "\Path.dat"
    If Dir(Archivo, vbHidden) <> "" Then SetAttr Archivo, vbNormal: Kill (Archivo)
    Archivo = Path & "\Huellas.dat"
    If Dir(Archivo, vbHidden) <> "" Then SetAttr Archivo, vbNormal: Kill (Archivo)
    Archivo = Path & "\HuellasDigitales.tmp"
    If Dir(Archivo) <> "" Then Kill (Archivo)
End Sub

Private Sub cmbEmpleado_Change()
    Deshabilitar False
    If cmbEmpleado.Text <> "" Then
        'Verificar si la cedula ya existe y carga el profesor
        If CargarProfesor = True Then
            strSql = " SELECT epl_baja " & _
                     " FROM empleado " & _
                     " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                     " AND emp_codigo='" & strEmpresa & "'"
            clsSql.Ejecutar strSql
            If clsSql.adorec_Def.RecordCount > 0 Then
                If CBool(clsSql.adorec_Def(0)) Then
                    Deshabilitar True
                End If
            End If
            btnAgregar.Caption = "&Modificar"
            Actualizar = True
            
        Else
            btnAgregar.Caption = "&Aceptar"
            Actualizar = False
        End If
    End If
End Sub



Private Sub txtClave_Change()
    Dim clave As String
    clave = txtConfirmaClave.Text
    If txtClave.Text <> clave Then
        txtConfirmaClave.ForeColor = vbRed
    Else
        txtConfirmaClave.ForeColor = vbDefault
    End If
End Sub

Private Sub txtClave_GotFocus()
    'SeleccionarContenido
End Sub

Private Sub txtConfirmaClave_Change()
    Dim clave As String
    clave = txtConfirmaClave.Text
    If txtClave.Text <> clave Then
        txtConfirmaClave.ForeColor = vbRed
    Else
        txtConfirmaClave.ForeColor = vbDefault
    End If
End Sub

Private Sub txtConfirmaClave_GotFocus()
    'SeleccionarContenido
End Sub

Private Function ComprobarClave() As Boolean
    Dim clave As String, confirmaclave As String
    clave = txtClave.Text
    confirmaclave = txtConfirmaClave.Text
    If clave = "" Then
        MsgBox "Ingrese una contraseña para el registro", vbCritical, "Contraseña"
        ComprobarClave = False
    ElseIf Len(txtClave.Text) <> 3 Then
        MsgBox "La contraseña debe tener 3 caracteres", vbCritical, "Contraseña"
        ComprobarClave = False
    ElseIf clave <> confirmaclave Then
        MsgBox "La contraseña no concuerda, confirme la contraseña correcta", vbCritical, "Contraseña"
        ComprobarClave = False
    Else
        ComprobarClave = True
   End If
End Function


Private Function GenerarClave() As Boolean
    Dim ci As String, clave As String
    clave = txtClave.Text
    If Actualizar = False Then
        
        strSql = " SELECT epl_codigo FROM empleado_huella WHERE epl_huella_clave='" & clave & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql
        
        If clsSql.adorec_Def.RecordCount = 0 Then
            codigo = clave
            GenerarClave = True
        Else
            codigo = ""
            GenerarClave = False
        End If
    Else
        
        codigo = clave
        GenerarClave = True
    End If
End Function


Private Function ComprobarDatos(Tipo As Integer) As Boolean
    If Tipo = 0 Then
        If cmbEmpleado.Text = "" Then
            MsgBox "Seleccione un Empleado", vbCritical, "Registro de Huella"
            ComprobarDatos = False
        Else
            ComprobarDatos = True
        End If
    Else
        If cmbEmpleado.Text = "" Then
            MsgBox "Seleccione un Empleado", vbCritical, "Registro de Huella"
            ComprobarDatos = False
        Else
            ComprobarDatos = True
        End If
    End If
End Function


Private Function CargarProfesor() As Boolean
    Dim huellas As Boolean
    
    strSql = " SELECT COALESCE(epl_codigo,'') as epl_codigo,COALESCE(epl_huella_registrado,'0') as epl_huella_registrado,COALESCE(epl_huella_clave,'') as epl_huella_clave,epl_huella_base,epl_huella_interno " & _
             " FROM empleado_huella " & _
             " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
             " AND emp_codigo='" & strEmpresa & "' "
    clsSql.Ejecutar strSql
    If clsSql.adorec_Def.RecordCount > 0 Then
    
        codigo = clsSql.adorec_Def("epl_huella_clave")
        
        
        
        If Interno = True Then
            Registrado = CBool(clsSql.adorec_Def("epl_huella_interno"))
            If CBool(clsSql.adorec_Def("epl_huella_interno")) = True Then
                btnAgregarHuella.Enabled = False
                btnBorrarHuella.Enabled = True
                btnVerificarHuella.Enabled = True
            Else
                btnAgregarHuella.Enabled = True
                btnBorrarHuella.Enabled = False
                btnVerificarHuella.Enabled = False
            End If
        Else
            Registrado = CBool(clsSql.adorec_Def("epl_huella_base"))
            If CBool(clsSql.adorec_Def("epl_huella_base")) = True Then
                btnAgregarHuella.Enabled = False
                btnBorrarHuella.Enabled = True
                btnVerificarHuella.Enabled = True
            Else
                btnAgregarHuella.Enabled = True
                btnBorrarHuella.Enabled = False
                btnVerificarHuella.Enabled = False
            End If
        End If
        txtClave.Text = clsSql.adorec_Def("epl_huella_clave")
        txtConfirmaClave.Text = clsSql.adorec_Def("epl_huella_clave")
        CargarProfesor = True
    Else
        btnAgregarHuella.Enabled = True
        CargarProfesor = False
    End If
    
End Function


Private Sub Modificar()
    Dim codigo, clave As String, i As Long
    
    clave = txtClave.Text
    
    If Interno = True Then
        strSql = " UPDATE empleado_huella SET " & _
                 " epl_huella_registrado='" & IIf(Registrado = True, 1, 0) & "'," & _
                 " epl_huella_interno='" & IIf(Registrado = True, 1, 0) & "'," & _
                 " epl_huella_clave='" & clave & "'," & _
                 " epl_huella_fechamod=CURRENT_TIMESTAMP," & _
                 " epl_huella_usumod='" & strUsuario & "' " & _
                 " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSql, "M"
    Else
        strSql = " UPDATE empleado_huella SET " & _
                 " epl_huella_registrado='" & IIf(Registrado = True, 1, 0) & "'," & _
                 " epl_huella_base='" & IIf(Registrado = True, 1, 0) & "'," & _
                 " epl_huella_clave='" & clave & "'," & _
                 " epl_huella_fechamod=CURRENT_TIMESTAMP," & _
                 " epl_huella_usumod='" & strUsuario & "' " & _
                 " WHERE epl_codigo='" & cmbEmpleado.BoundText & "' " & _
                 " AND emp_codigo='" & strEmpresa & "'"
        clsSql.Ejecutar strSql, "M"
    End If
End Sub

Private Sub BorrarHuella(cod As String)
    If Interno = True Then
        strSql = " UPDATE empleado_huella " & _
                 " SET epl_huella_registrado='0', " & _
                 " epl_huella_interno='0' " & _
                 " WHERE epl_codigo='" & cod & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql, "M"
    Else
        strSql = " UPDATE empleado_huella " & _
                 " SET epl_huella_registrado='0', " & _
                 " epl_huella_base='0' " & _
                 " WHERE epl_codigo='" & cod & "' AND emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql, "M"
    End If
    Registrado = False
End Sub

Private Sub Deshabilitar(Tipo As Boolean)
    Dim Valor As Boolean
    Valor = Not Tipo
    If Tipo = True Then
        btnAgregar.Enabled = Valor
        fraContenedor.Enabled = Valor
    Else
        btnAgregar.Enabled = Valor
        fraContenedor.Enabled = Valor
    End If
End Sub

Private Sub GuardaHuella()
    Dim rs As New ADODB.Recordset
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    On Error GoTo mistake
    
    If Interno = True Then
        strSql = " DELETE FROM huella WHERE emp_codigo='" & strEmpresa & "' "
        clsSql.Ejecutar strSql, "M"
        strSql = " REPAIR TABLE huella "
        clsSql.Ejecutar strSql, "M"
        clsSql.Ejecutar strSql
        strSql = " OPTIMIZE TABLE huella "
        clsSql.Ejecutar strSql, "M"
        clsSql.Ejecutar strSql
        rs.Open "SELECT * FROM huella WHERE emp_codigo='" & strEmpresa & "' ", AdoConn, adOpenKeyset, adLockOptimistic
    Else
        strSql = " DELETE FROM huella "
        clsAux.Ejecutar strSql, "M"
        strSql = " REPAIR TABLE huella "
        clsAux.Ejecutar strSql, "M"
        clsAux.Ejecutar strSql
        strSql = " OPTIMIZE TABLE huella "
        clsAux.Ejecutar strSql, "M"
        clsAux.Ejecutar strSql
        
        rs.Open "SELECT * FROM huella ", conSlave, adOpenKeyset, adLockOptimistic
    End If
        
    rs.AddNew
    mystream.Open
    mystream.LoadFromFile strPathHuella & "HuellasDigitales.tmp"
    rs!Archivo = mystream.Read
    If Interno = True Then
        rs!emp_codigo = strEmpresa
    End If
    rs.Update
    mystream.Close
    rs.Close
    Exit Sub
mistake:
    MsgBox "No se pudo guardar la huella, inténtelo nuevamente", vbInformation, "Huella Digital"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbKeyTab
    End If
End Sub


Private Function CargarHuella() As Boolean
    On Error GoTo errorCarga
    Dim rs As New ADODB.Recordset
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    
    Archivo = Path & "\HuellasDigitales.tmp"
    If Dir(Archivo) <> "" Then Kill (Archivo)
    
    If Interno = True Then
        rs.Open "SELECT * FROM huella WHERE emp_codigo='" & strEmpresa & "' ", AdoConn
    Else
        rs.Open "SELECT * FROM huella ", conSlave
    End If
    
    mystream.Open
        
    If rs.RecordCount = 0 Then
        'MsgBox "No esta Cargada la Aplicaión"
        num = FreeFile
        Archivo = Path & "\HuellasDigitales.tmp"
        
        
        If Dir(Archivo) <> "" Then
            Open Archivo For Output As #num
            
            Close #num
        End If
        
        CargarHuella = False
    Else
        mystream.Write rs!Archivo
        mystream.SaveToFile "HuellasDigitales.tmp", adSaveCreateOverWrite
        CargarHuella = True
    End If
    mystream.Close
    rs.Close
    
    Exit Function
    
errorCarga:
    MsgBox "Ocurrió un error en la carga de la Base de Huellas Digitales", vbCritical + vbOKOnly, "Base de Datos"
    CargarHuella = False
End Function


